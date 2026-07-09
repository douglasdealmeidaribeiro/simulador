const fs = require('fs');
const { HyperFormula } = require('hyperformula');

global.window = {};
eval(fs.readFileSync('assets/workbook-model.js', 'utf8'));

const model = window.WORKBOOK_MODEL;

function colToNum(col) {
  return [...col].reduce((acc, char) => acc * 26 + char.charCodeAt(0) - 64, 0);
}

function refToAddress(ref) {
  const match = ref.match(/^([A-Z]+)(\d+)$/);
  return { row: Number(match[2]) - 1, col: colToNum(match[1]) - 1 };
}

function numToCol(num) {
  let out = '';
  for (num += 1; num; num = Math.floor((num - 1) / 26)) {
    out = String.fromCharCode(65 + ((num - 1) % 26)) + out;
  }
  return out;
}

function addressToRef(address) {
  return `${numToCol(address.col)}${address.row + 1}`;
}

function valueForCompare(value) {
  if (value && typeof value === 'object' && value.value) {
    return value.value;
  }
  return value;
}

const sheets = {};
for (const sheet of model.sheets) {
  const rows = Array.from({ length: sheet.maxRow }, () => []);
  for (const [ref, value] of Object.entries(sheet.cells)) {
    const address = refToAddress(ref);
    rows[address.row][address.col] = value;
  }
  sheets[sheet.name] = rows;
}

console.time('build');
const hf = HyperFormula.buildFromSheets(sheets, {
  licenseKey: 'gpl-v3',
  useArrayArithmetic: true,
});
console.timeEnd('build');

function inspectSheet(sheetName, rowStart, rowEnd, colStart, colEnd, limit = 40) {
  const sheet = model.sheets.find((candidate) => candidate.name === sheetName);
  const sheetId = hf.getSheetId(sheetName);
  const diffs = [];

  for (let row = rowStart; row <= rowEnd; row += 1) {
    for (let col = colToNum(colStart); col <= colToNum(colEnd); col += 1) {
      const ref = `${numToCol(col - 1)}${row}`;
      if (!(ref in (sheet.cache || {}))) {
        continue;
      }
      const address = refToAddress(ref);
      const calculated = valueForCompare(hf.getCellValue({
        sheet: sheetId,
        row: address.row,
        col: address.col,
      }));
      const cached = sheet.cache[ref];
      if (typeof calculated === 'number' && typeof cached === 'number') {
        const abs = Math.abs(calculated - cached);
        const rel = abs / Math.max(1, Math.abs(cached));
        if (abs > 1e-6 && rel > 1e-8) {
          diffs.push({ ref, calculated, cached, abs, rel });
        }
      } else if (calculated !== cached) {
        diffs.push({ ref, calculated, cached, abs: Number.POSITIVE_INFINITY, rel: Number.POSITIVE_INFINITY });
      }
    }
  }

  diffs.sort((a, b) => b.abs - a.abs);
  for (const diff of diffs.slice(0, limit)) {
    const address = refToAddress(diff.ref);
    const formula = hf.getCellFormula({ sheet: sheetId, row: address.row, col: address.col });
    console.log(`${sheetName}!${diff.ref}`, {
      calculated: diff.calculated,
      cached: diff.cached,
      abs: diff.abs,
      rel: diff.rel,
      formula,
    });
  }
}

inspectSheet('DFs', 1, 120, 'H', 'AK');
inspectSheet('Inputs mensais', 1, 160, 'H', 'AK');
inspectSheet('Orçamento (Mensal)', 1, 120, 'Q', 'AK');
