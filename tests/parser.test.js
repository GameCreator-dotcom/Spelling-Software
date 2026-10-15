const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const test = require('node:test');

const context = {};
vm.createContext(context);
vm.runInContext(fs.readFileSync(new URL('../Code.gs', `file://${__dirname}/`), 'utf8'), context);

test('parses Collegiate definitions, part of speech, etymology, and supplied examples', () => {
  const entry = {
    meta: { id: 'voluminous' },
    fl: 'adjective',
    shortdef: ['having great {a_link|volume} or bulk : large', 'numerous'],
    et: [['text', 'Late Latin {it}voluminosus{/it}, from Latin {it}volumen{/it}']],
    suppl: { examples: [{ t: 'a {it}voluminous{/it} collection' }] }
  };
  assert.deepEqual(JSON.parse(JSON.stringify(context.parseDictionaryResponse_([entry], 'voluminous'))), {
    definition: 'having great volume or bulk : large; numerous',
    partOfSpeech: 'adjective',
    origin: 'Late Latin voluminosus, from Latin volumen',
    example: 'a voluminous collection'
  });
});

test('selects an exact headword and extracts a definition example when no supplied example exists', () => {
  const response = [
    { meta: { id: 'benevolent:1' }, fl: 'adjective', shortdef: ['well meaning'], def: [{ sseq: [[['sense', { dt: [['vis', [{ t: 'a {wi}benevolent{/wi} act' }]]] }]]] }] },
    { meta: { id: 'benevolentia' }, fl: 'noun', shortdef: ['goodwill'] }
  ];
  assert.equal(context.parseDictionaryResponse_(response, 'benevolent').example, 'a benevolent act');
  assert.equal(context.parseDictionaryResponse_(response, 'benevolent').partOfSpeech, 'adjective');
});

test('reports missing words and leaves unavailable fields blank', () => {
  assert.equal(context.parseDictionaryResponse_(['suggestion'], 'unknown').definition, 'Not found in Merriam-Webster.');
  const parsed = context.parseDictionaryResponse_([{ meta: { id: 'word' }, shortdef: ['a definition'] }], 'word');
  assert.equal(parsed.definition, 'a definition');
  assert.equal(parsed.partOfSpeech, '');
  assert.equal(parsed.origin, '');
  assert.equal(parsed.example, '');
});

test('fills multiple rows while preserving existing values and formula cells', () => {
  const values = [
    ['first', '', '', '', ''],
    ['second', 'User definition', '', '', ''],
    ['third', '', '', '', '']
  ];
  const formulas = [
    ['', '', '', '', ''],
    ['', '', '=CUSTOM()', '', ''],
    ['', '', '', '', '']
  ];
  const writes = [];
  const sheet = {
    getRange(row, column, rowCount, columnCount) {
      return {
        getValues: () => values.slice(row - 2, row - 2 + rowCount).map((item) => item.slice(0, columnCount)),
        getFormulas: () => formulas.slice(row - 2, row - 2 + rowCount).map((item) => item.slice(0, columnCount)),
        setValues(nextValues) {
          writes.push({ row, column, nextValues });
          nextValues.forEach((line, index) => {
            values[row - 2 + index][column - 1] = line[0];
          });
        }
      };
    }
  };
  context.LockService = { getDocumentLock: () => ({ tryLock: () => true, releaseLock() {} }) };
  context.lookupWords_ = (words) => words.map((word) => ({
    definition: `${word} definition`, partOfSpeech: 'noun', origin: '', example: ''
  }));
  context.processVocabularyRows_(sheet, 2, 3, false);

  assert.equal(values[0][1], 'first definition');
  assert.equal(values[1][1], 'User definition');
  assert.equal(values[1][2], '');
  assert.equal(values[2][1], 'third definition');
  assert.ok(writes.every((write) => write.row !== 3 || write.column !== 3));
  assert.equal(writes.length, 6);
});