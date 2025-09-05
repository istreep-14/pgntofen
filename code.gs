function run() {
  const grouped = buildGroupedOpenings();
  renderGroupedOpeningsToSheet(grouped);
}

// Fetch and group A–E TSVs
function buildGroupedOpenings() {
  const files = ['a.tsv', 'b.tsv', 'c.tsv', 'd.tsv', 'e.tsv'];
  const baseUrl = 'https://raw.githubusercontent.com/lichess-org/chess-openings/master/';
  const openings = [];

  files.forEach(file => {
    const text = UrlFetchApp.fetch(baseUrl + file).getContentText();
    const lines = text.split(/\r?\n/).filter(Boolean);
    const first = lines[0] || '';
    const hasHeader = first.toLowerCase().replace(/^\uFEFF/, '').startsWith('eco\t');
    const startIdx = hasHeader ? 1 : 0;

    for (let i = startIdx; i < lines.length; i++) {
      const parts = lines[i].split('\t');
      if (parts.length < 3) continue;
      const eco = (parts[0] || '').trim();
      const name = (parts[1] || '').trim();
      const pgn = (parts[2] || '').trim();
      if (!eco || !name || !pgn) continue;

      const parsed = parseName(name);
      openings.push({
        eco, name, pgn,
        family: parsed.family,
        variation: parsed.variation,
        subvariations: parsed.subvariations
      });
    }
  });

  // Build nested: family -> variation -> subvariation -> [entries]
  const grouped = {};
  for (const op of openings) {
    const family = op.family || '—';
    const variation = op.variation || '—';
    const subvars = op.subvariations.length ? op.subvariations : ['—'];

    if (!grouped[family]) grouped[family] = {};
    if (!grouped[family][variation]) grouped[family][variation] = {};

    // Place the opening under each subvariation it lists
    for (const sv of subvars) {
      if (!grouped[family][variation][sv]) grouped[family][variation][sv] = [];
      grouped[family][variation][sv].push({
        eco: op.eco,
        name: op.name,
        pgn: op.pgn
      });
    }
  }
  return grouped;
}

// Robust parse of "Family: Variation, Subvar, Subvar2"
function parseName(name) {
  const idx = name.indexOf(':');
  if (idx === -1) {
    return { family: name.trim(), variation: '', subvariations: [] };
  }
  const family = name.slice(0, idx).trim();
  const tail = name.slice(idx + 1).trim();
  if (!tail) return { family, variation: '', subvariations: [] };

  const parts = tail.split(',').map(s => s.trim()).filter(Boolean);
  const variation = parts.length ? parts[0] : '';
  const subvariations = parts.length > 1 ? parts.slice(1) : [];
  return { family, variation, subvariations };
}

// Write grouped structure into a sheet with outline groups
function renderGroupedOpeningsToSheet(grouped) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const title = 'Grouped Openings';
  const sheet = ss.getSheetByName(title) || ss.insertSheet(title);
  sheet.clear();

  const header = ['Level', 'Family', 'Variation', 'Subvariation', 'ECO', 'Name', 'PGN'];
  sheet.getRange(1, 1, 1, header.length).setValues([header]);
  sheet.setFrozenRows(1);

  let row = 2;
  const lastCol = header.length;

  const sortKeys = obj => Object.keys(obj).sort((a, b) =>
    a.localeCompare(b, undefined, { sensitivity: 'base' })
  );

  for (const family of sortKeys(grouped)) {
    const familyHeaderRow = row++;
    sheet.getRange(familyHeaderRow, 1, 1, lastCol).setValues([['Family', family, '', '', '', '', '']]);

    const variationBlockStart = row;

    for (const variation of sortKeys(grouped[family])) {
      const variationHeaderRow = row++;
      sheet.getRange(variationHeaderRow, 1, 1, lastCol).setValues([['Variation', family, variation, '', '', '', '']]);

      const subvarBlockStart = row;

      for (const subvar of sortKeys(grouped[family][variation])) {
        const subvarHeaderRow = row++;
        sheet.getRange(subvarHeaderRow, 1, 1, lastCol).setValues([['Subvariation', family, variation, subvar, '', '', '']]);

        const entries = grouped[family][variation][subvar];
        if (entries.length) {
          const dataRows = entries.map(e => ['Entry', family, variation, subvar, e.eco, e.name, e.pgn]);
          sheet.getRange(row, 1, dataRows.length, lastCol).setValues(dataRows);
          sheet.getRange(subvarHeaderRow + 1, 1, dataRows.length, lastCol).shiftRowGroupDepth(1);
          row += dataRows.length;
        }
      }

      if (row > subvarBlockStart) {
        sheet.getRange(subvarBlockStart, 1, row - subvarBlockStart, lastCol).shiftRowGroupDepth(1);
      }
    }

    if (row > variationBlockStart) {
      sheet.getRange(variationBlockStart, 1, row - variationBlockStart, lastCol).shiftRowGroupDepth(1);
    }
  }

  for (let c = 1; c <= lastCol; c++) sheet.autoResizeColumn(c);
}

/**
 * Incremental, resume-safe import for Lichess chess-openings A–E TSVs.
 * Writes an append-only sheet with columns:
 * Family | Variation | Subvariation | ECO | Name | PGN | SourceFile | DataLine | Key
 *
 * Run resumeImport() repeatedly (or via trigger); it picks up where it left off.
 */

const CONFIG = {
  baseUrl: 'https://raw.githubusercontent.com/lichess-org/chess-openings/master/',
  files: ['a.tsv', 'b.tsv', 'c.tsv', 'd.tsv', 'e.tsv'],
  targetSheetName: 'Openings Normalized',
  batchInputLines: 400, // adjust if needed
  headerRow: ['Family', 'Variation', 'Subvariation', 'ECO', 'Name', 'PGN', 'SourceFile', 'DataLine', 'Key'],
  propFileIdx: 'CO_FILE_IDX',
  propDataIdx: 'CO_DATA_IDX'
};

function resumeImport() {
  const props = PropertiesService.getScriptProperties();
  let fileIdx = parseInt(props.getProperty(CONFIG.propFileIdx) || '0', 10);
  let dataIdx = parseInt(props.getProperty(CONFIG.propDataIdx) || '0', 10);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateTargetSheet_(ss);
  const existingKeySet = getExistingKeySet_(sheet);

  let remainingLines = CONFIG.batchInputLines;
  const rowsToAppend = [];

  for (; fileIdx < CONFIG.files.length && remainingLines > 0; fileIdx++) {
    const fileName = CONFIG.files[fileIdx];
    const { dataLines } = fetchDataLines_(fileName);

    let i = (fileIdx === parseInt(props.getProperty(CONFIG.propFileIdx) || '0', 10)) ? dataIdx : 0;

    while (i < dataLines.length && remainingLines > 0) {
      const raw = dataLines[i];
      i++;
      remainingLines--;

      const parts = raw.split('\t');
      if (parts.length < 3) continue;

      const eco = (parts[0] || '').trim();
      const name = (parts[1] || '').trim();
      const pgn = (parts[2] || '').trim();
      if (!eco || !name || !pgn) continue;

      const { family, variation, subvariations } = parseName_(name);
      const subvars = (subvariations && subvariations.length) ? subvariations : ['—'];

      for (const sv of subvars) {
        const row = [family || '—', variation || '—', sv, eco, name, pgn, fileName, i, ''];
        row[8] = makeKey_(row);
        if (!existingKeySet.has(row[8])) {
          rowsToAppend.push(row);
          existingKeySet.add(row[8]);
        }
      }
    }

    if (i < dataLines.length) {
      dataIdx = i;
      break;
    } else {
      dataIdx = 0;
    }
  }

  if (rowsToAppend.length) {
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rowsToAppend.length, CONFIG.headerRow.length).setValues(rowsToAppend);
  }

  PropertiesService.getScriptProperties().setProperty(CONFIG.propFileIdx, String(fileIdx));
  PropertiesService.getScriptProperties().setProperty(CONFIG.propDataIdx, String(dataIdx));
}

function isComplete() {
  const props = PropertiesService.getScriptProperties();
  const fileIdx = parseInt(props.getProperty(CONFIG.propFileIdx) || '0', 10);
  const dataIdx = parseInt(props.getProperty(CONFIG.propDataIdx) || '0', 10);
  if (fileIdx >= CONFIG.files.length) return true;
  const { dataLines } = fetchDataLines_(CONFIG.files[fileIdx]);
  return (fileIdx >= CONFIG.files.length) || (dataIdx >= dataLines.length && fileIdx === CONFIG.files.length - 1);
}

function resetProgress(clearSheet) {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty(CONFIG.propFileIdx);
  props.deleteProperty(CONFIG.propDataIdx);

  if (clearSheet) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.targetSheetName);
    if (sheet) {
      sheet.clear();
      sheet.getRange(1, 1, 1, CONFIG.headerRow.length).setValues([CONFIG.headerRow]);
      sheet.setFrozenRows(1);
    }
  }
}

/* Optional: auto-continue via triggers */
function createMinuteTrigger() {
  ScriptApp.newTrigger('resumeImport').timeBased().everyMinutes(1).create();
}

function deleteAllTriggers() {
  ScriptApp.getProjectTriggers().forEach(ScriptApp.deleteTrigger);
}

/* ---------- Helpers ---------- */

function getOrCreateTargetSheet_(ss) {
  let sheet = ss.getSheetByName(CONFIG.targetSheetName);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.targetSheetName);
    sheet.getRange(1, 1, 1, CONFIG.headerRow.length).setValues([CONFIG.headerRow]);
    sheet.setFrozenRows(1);
  } else if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, CONFIG.headerRow.length).setValues([CONFIG.headerRow]);
    sheet.setFrozenRows(1);
  } else {
    const firstRow = sheet.getRange(1, 1, 1, CONFIG.headerRow.length).getValues()[0];
    if (firstRow.join('\t') !== CONFIG.headerRow.join('\t')) {
      sheet.insertRows(1);
      sheet.getRange(1, 1, 1, CONFIG.headerRow.length).setValues([CONFIG.headerRow]);
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
}

function getExistingKeySet_(sheet) {
  const lastRow = sheet.getLastRow();
  const keyColIndex = CONFIG.headerRow.indexOf('Key') + 1;
  const keys = new Set();
  if (lastRow <= 1) return keys;
  const values = sheet.getRange(2, keyColIndex, lastRow - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    const k = (values[i][0] || '').toString();
    if (k) keys.add(k);
  }
  return keys;
}

function fetchDataLines_(fileName) {
  const url = CONFIG.baseUrl + fileName;
  const text = UrlFetchApp.fetch(url).getContentText();
  const lines = text.split(/\r?\n/).filter(line => line.length > 0);
  const first = (lines[0] || '').replace(/^\uFEFF/, '');
  const hasHeader = first.toLowerCase().startsWith('eco\t');
  const startIdx = hasHeader ? 1 : 0;
  return { dataLines: lines.slice(startIdx) };
}

function parseName_(name) {
  const idx = name.indexOf(':');
  if (idx === -1) return { family: name.trim(), variation: '', subvariations: [] };
  const family = name.slice(0, idx).trim();
  const tail = name.slice(idx + 1).trim();
  if (!tail) return { family, variation: '', subvariations: [] };
  const parts = tail.split(',').map(s => s.trim()).filter(Boolean);
  const variation = parts.length ? parts[0] : '';
  const subvariations = parts.length > 1 ? parts.slice(1) : [];
  return { family, variation, subvariations };
}

function makeKey_(row) {
  // row = [Family, Variation, Subvariation, ECO, Name, PGN, SourceFile, DataLine, Key]
  const family = row[0] || '';
  const variation = row[1] || '';
  const subvar = row[2] || '';
  const eco = row[3] || '';
  const name = row[4] || '';
  const pgn = row[5] || '';
  return [eco, name, pgn, family, variation, subvar].join('|');
}
