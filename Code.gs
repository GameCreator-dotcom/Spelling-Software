const VOCABULARY_HEADERS = [
  'Word',
  'Definition',
  'Part of Speech',
  'Language of Origin',
  'Word in a Sentence'
];
const MAX_ROWS_PER_RUN = 50;
const CACHE_TTL_SECONDS = 21600;
const API_BASE_URL = 'https://www.dictionaryapi.com/api/v3/references/collegiate/json/';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Merriam-Webster Vocabulary')
    .addItem('Set Up Vocabulary Sheet', 'setUpVocabularySheet')
    .addItem('Look Up Selected Words', 'lookUpSelectedWords')
    .addItem('Refresh Selected Words', 'refreshSelectedWords')
    .addItem('Help', 'showVocabularyHelp')
    .addToUi();
}

function setUpVocabularySheet() {
  const sheet = SpreadsheetApp.getActiveSheet();
  sheet.getRange(1, 1, 1, VOCABULARY_HEADERS.length).setValues([VOCABULARY_HEADERS]);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, VOCABULARY_HEADERS.length)
    .setFontWeight('bold')
    .setBackground('#d9ead3');
  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 360);
  sheet.setColumnWidth(3, 150);
  sheet.setColumnWidth(4, 300);
  sheet.setColumnWidth(5, 360);
  if (!sheet.getFilter()) {
    sheet.getRange(1, 1, Math.max(sheet.getMaxRows(), 2), VOCABULARY_HEADERS.length).createFilter();
  }

  const spreadsheet = SpreadsheetApp.getActive();
  const installed = ScriptApp.getProjectTriggers().some(function (trigger) {
    return trigger.getHandlerFunction() === 'handleVocabularyEdit' &&
      trigger.getTriggerSourceId() === spreadsheet.getId();
  });
  if (!installed) {
    ScriptApp.newTrigger('handleVocabularyEdit')
      .forSpreadsheet(spreadsheet)
      .onEdit()
      .create();
  }

  SpreadsheetApp.getUi().alert('Vocabulary sheet is ready. Enter words in column A.');
}

function handleVocabularyEdit(event) {
  if (!event || !event.range) return;
  const range = event.range;
  if (range.getColumn() > 1 || range.getLastColumn() < 1 || range.getLastRow() < 2) return;
  const firstRow = Math.max(2, range.getRow());
  processVocabularyRows_(range.getSheet(), firstRow, range.getLastRow() - firstRow + 1, false);
}

function lookUpSelectedWords() {
  processSelection_(false);
}

function refreshSelectedWords() {
  processSelection_(true);
}

function processSelection_(forceRefresh) {
  const range = SpreadsheetApp.getActiveRange();
  if (!range) return;
  const sheet = range.getSheet();
  const firstRow = Math.max(2, range.getRow());
  const lastRow = range.getLastRow();
  if (lastRow < firstRow) return;
  processVocabularyRows_(sheet, firstRow, lastRow - firstRow + 1, forceRefresh);
}

function processVocabularyRows_(sheet, firstRow, rowCount, forceRefresh) {
  if (rowCount <= 0) return;
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(10000)) return;
  try {
    const count = Math.min(rowCount, MAX_ROWS_PER_RUN);
    const range = sheet.getRange(firstRow, 1, count, VOCABULARY_HEADERS.length);
    const rows = range.getValues();
    const formulas = range.getFormulas();
    const pending = [];
    rows.forEach(function (row, index) {
      const word = String(row[0] || '').trim();
      if (!word) return;
      const hasBlankField = row.slice(1).some(function (value, fieldIndex) {
        return (value === '' || value === null) && !formulas[index][fieldIndex + 1];
      });
      if (forceRefresh || hasBlankField) pending.push({ index: index, word: word });
    });
    if (!pending.length) return;

    const results = lookupWords_(pending.map(function (item) { return item.word; }));
    const updatesByColumn = [[], [], [], []];
    pending.forEach(function (item, resultIndex) {
      const result = results[resultIndex];
      const row = rows[item.index];
      const fields = [result.definition, result.partOfSpeech, result.origin, result.example];
      fields.forEach(function (value, fieldIndex) {
        const blank = row[fieldIndex + 1] === '' || row[fieldIndex + 1] === null;
        const hasFormula = Boolean(formulas[item.index][fieldIndex + 1]);
        if (forceRefresh || (blank && !hasFormula)) {
          row[fieldIndex + 1] = value;
          updatesByColumn[fieldIndex].push({ row: firstRow + item.index, value: value });
        }
      });
    });
    updatesByColumn.forEach(function (updates, columnIndex) {
      let offset = 0;
      while (offset < updates.length) {
        let end = offset + 1;
        while (end < updates.length && updates[end].row === updates[end - 1].row + 1) end++;
        sheet.getRange(updates[offset].row, columnIndex + 2, end - offset, 1)
          .setValues(updates.slice(offset, end).map(function (update) { return [update.value]; }));
        offset = end;
      }
    });
  } finally {
    lock.releaseLock();
  }
}

function lookupWords_(words) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('MW_COLLEGIATE_API_KEY');
  if (!apiKey) {
    return words.map(function () {
      return emptyResult_('API key not configured. Add MW_COLLEGIATE_API_KEY in Apps Script Script Properties.');
    });
  }
  const cache = CacheService.getScriptCache();
  const results = new Array(words.length);
  const uncached = [];
  words.forEach(function (word, index) {
    const key = cacheKey_(word);
    const cached = cache.get(key);
    if (cached) results[index] = JSON.parse(cached);
    else uncached.push({ word: word, index: index, cacheKey: key });
  });

  for (let offset = 0; offset < uncached.length; offset += 10) {
    const batch = uncached.slice(offset, offset + 10);
    const requests = batch.map(function (item) {
      return {
        url: API_BASE_URL + encodeURIComponent(item.word) + '?key=' + encodeURIComponent(apiKey),
        muteHttpExceptions: true
      };
    });
    let responses;
    try {
      responses = UrlFetchApp.fetchAll(requests);
    } catch (error) {
      batch.forEach(function (item) {
        results[item.index] = emptyResult_('Merriam-Webster request failed. Check API access and try again.');
      });
      continue;
    }
    responses.forEach(function (response, index) {
      const item = batch[index];
      const code = response.getResponseCode();
      if (code < 200 || code >= 300) {
        results[item.index] = { definition: 'Merriam-Webster lookup failed (HTTP ' + code + ').', partOfSpeech: '', origin: '', example: '' };
        return;
      }
      let payload;
      try {
        payload = JSON.parse(response.getContentText());
      } catch (error) {
        results[item.index] = { definition: 'Merriam-Webster returned an unreadable response.', partOfSpeech: '', origin: '', example: '' };
        return;
      }
      const parsed = parseDictionaryResponse_(payload, item.word);
      results[item.index] = parsed;
      if (parsed.definition !== 'Not found in Merriam-Webster.') {
        cache.put(item.cacheKey, JSON.stringify(parsed), CACHE_TTL_SECONDS);
      }
    });
    if (offset + 10 < uncached.length) Utilities.sleep(100);
  }
  return results;
}

function parseDictionaryResponse_(payload, requestedWord) {
  if (!Array.isArray(payload)) return emptyResult_('Merriam-Webster returned an unexpected response.');
  const entries = payload.filter(function (entry) { return entry && typeof entry === 'object' && entry.meta; });
  if (!entries.length) return emptyResult_('Not found in Merriam-Webster.');

  const normalizedRequest = normalizeHeadword_(requestedWord);
  const entry = entries.find(function (candidate) {
    const id = normalizeHeadword_(String(candidate.meta.id || '').replace(/:\d+$/, ''));
    return id === normalizedRequest;
  }) || entries[0];
  const definition = Array.isArray(entry.shortdef)
    ? entry.shortdef.slice(0, 3).map(cleanDictionaryText_).filter(Boolean).join('; ')
    : '';
  const origin = extractEtymology_(entry.et);
  const example = extractExamples_(entry);
  return {
    definition: definition || 'Not supplied by Merriam-Webster.',
    partOfSpeech: cleanDictionaryText_(entry.fl || ''),
    origin: origin || '',
    example: example || ''
  };
}

function extractEtymology_(etymology) {
  if (!Array.isArray(etymology)) return '';
  const texts = etymology.filter(function (item) {
    return Array.isArray(item) && item[0] === 'text' && typeof item[1] === 'string';
  }).map(function (item) { return cleanDictionaryText_(item[1]); }).filter(Boolean);
  return texts.join(' ');
}

function extractExamples_(entry) {
  const supplied = entry.suppl && Array.isArray(entry.suppl.examples) ? entry.suppl.examples : [];
  for (let index = 0; index < supplied.length; index++) {
    const text = cleanDictionaryText_(supplied[index] && supplied[index].t || '');
    if (text) return text;
  }
  const definitions = Array.isArray(entry.def) ? entry.def : [];
  for (let defIndex = 0; defIndex < definitions.length; defIndex++) {
    const senses = definitions[defIndex].sseq || [];
    const example = findVisExample_(senses);
    if (example) return example;
  }
  return '';
}

function findVisExample_(value) {
  if (!value || typeof value !== 'object') return '';
  if (Array.isArray(value)) {
    if (value[0] === 'vis' && Array.isArray(value[1])) {
      for (let index = 0; index < value[1].length; index++) {
        const text = cleanDictionaryText_(value[1][index] && value[1][index].t || '');
        if (text) return text;
      }
    }
    for (let index = 0; index < value.length; index++) {
      const found = findVisExample_(value[index]);
      if (found) return found;
    }
  } else {
    const keys = Object.keys(value);
    for (let index = 0; index < keys.length; index++) {
      const found = findVisExample_(value[keys[index]]);
      if (found) return found;
    }
  }
  return '';
}

function cleanDictionaryText_(value) {
  return String(value || '')
    .replace(/\{bc\}/g, ': ')
    .replace(/\{ldquo\}|\{rdquo\}/g, '"')
    .replace(/\{([a-z_]+)\|([^{}|]+)(?:\|[^{}]*)?\}/g, '$2')
    .replace(/\{\/?[a-z_]+\}/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function normalizeHeadword_(word) {
  return String(word || '').toLowerCase().replace(/[\s*]/g, '').replace(/:\d+$/, '');
}

function cacheKey_(word) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, normalizeHeadword_(word));
  const hex = digest.map(function (byte) {
    return ('0' + ((byte + 256) % 256).toString(16)).slice(-2);
  }).join('');
  return 'mw_collegiate_' + hex;
}

function emptyResult_(message) {
  return { definition: message, partOfSpeech: '', origin: '', example: '' };
}

function showVocabularyHelp() {
  SpreadsheetApp.getUi().alert(
    'Enter words in column A. Automatic lookups fill blank cells; use Refresh Selected Words to replace existing results. ' +
    'Configure MW_COLLEGIATE_API_KEY in Apps Script Project Settings > Script Properties.'
  );
}