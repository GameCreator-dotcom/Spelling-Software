# Merriam-Webster Vocabulary Assistant

A Google Sheets-bound Apps Script that fills vocabulary details from the official Merriam-Webster Collegiate Dictionary API. It does not use AI or another dictionary.

## Install

1. Create or open a Google Sheet and open **Extensions > Apps Script**.
2. Replace the starter script with `Code.gs` from this project and add the `appsscript.json` manifest from **Project Settings > Show "appsscript.json" manifest file**.
3. In Apps Script **Project Settings > Script Properties**, add `MW_COLLEGIATE_API_KEY` with your Collegiate Dictionary API key. The key is kept in Script Properties, not in the source.
4. Save and reload the Sheet. Choose **Merriam-Webster Vocabulary > Set Up Vocabulary Sheet** and approve Google's requested Sheets, external-request, and trigger permissions.

The setup action creates the five headers and installs an authorized edit trigger. The first use of URL fetching and trigger installation requires Google authorization. No Apps Script project can be deployed into your Google account from this local workspace.

## Use

Enter a word in column A to fill blank cells in columns B-E. Use **Look Up Selected Words** to fill blanks for selected rows, or **Refresh Selected Words** to replace all four dictionary fields with the current Merriam-Webster result. Missing Merriam-Webster fields stay blank; an unrecognized word is labeled in the Definition column. The lookup consistently uses the first matching entry and joins up to three short definitions.

The script caches successful lookups for six hours and limits each action to 50 rows. Only the Collegiate API is used; the School API key is not needed for this implementation.

## Local checks

Run `npm test` with Node.js installed. These tests exercise the response parser using Merriam-Webster-shaped fixtures; live requests require a valid API key and Google authorization.
