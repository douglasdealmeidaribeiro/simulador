self.window = self;
importScripts('assets/vendor/hyperformula.full.min.js');
importScripts('assets/workbook-model.js');

const { HyperFormula } = self.HyperFormula;
const model = self.WORKBOOK_MODEL;
let hf;
let sheetIds;

function colToNum(col) {
  return [...col].reduce((acc, char) => acc * 26 + char.charCodeAt(0) - 64, 0);
}

function refToAddress(ref, sheet) {
  const match = ref.match(/^([A-Z]+)(\d+)$/);
  return { sheet, row: Number(match[2]) - 1, col: colToNum(match[1]) - 1 };
}

function valueOf(value) {
  if (value && typeof value === 'object' && value.value) {
    throw new Error(value.message || value.value);
  }
  return value;
}

function get(sheet, ref) {
  return valueOf(hf.getCellValue(refToAddress(ref, sheet)));
}

function set(sheet, ref, value) {
  hf.setCellContents(refToAddress(ref, sheet), [[value]]);
}

function buildSheets() {
  const sheets = {};
  for (const sheet of model.sheets) {
    const rows = Array.from({ length: sheet.maxRow }, () => []);
    for (const [ref, value] of Object.entries(sheet.cells)) {
      const address = refToAddress(ref, 0);
      rows[address.row][address.col] = value;
    }
    sheets[sheet.name] = rows;
  }
  return sheets;
}

function currentResults() {
  const annual = get(sheetIds.simulador, 'G3');
  return {
    annual,
    monthly: get(sheetIds.simulador, 'G4'),
    target: get(sheetIds.controle, 'M13'),
  };
}

function evaluatePrice(price) {
  set(sheetIds.controle, 'M3', price);
  return Number(get(sheetIds.controle, 'M13'));
}

function goalSeek(goal = 0.1) {
  let low = 0;
  let high = 15000;
  let fLow = evaluatePrice(low) - goal;
  let fHigh = evaluatePrice(high) - goal;

  let guard = 0;
  while (fLow * fHigh > 0 && guard < 30) {
    high *= 1.6;
    fHigh = evaluatePrice(high) - goal;
    guard += 1;
  }

  if (fLow * fHigh > 0) {
    throw new Error('Não foi possível encontrar intervalo para a meta do VPL.');
  }

  let mid = high;
  let fMid = fHigh;
  let iterations = 0;

  for (; iterations < 80; iterations += 1) {
    mid = (low + high) / 2;
    fMid = evaluatePrice(mid) - goal;
    if (Math.abs(fMid) < 0.000001 || Math.abs(high - low) < 0.000001) {
      break;
    }
    if (fLow * fMid <= 0) {
      high = mid;
      fHigh = fMid;
    } else {
      low = mid;
      fLow = fMid;
    }
  }

  set(sheetIds.controle, 'M3', mid);
  return { price: mid, iterations, residual: fMid };
}

function applyUpdates(updates) {
  for (const [ref, value] of Object.entries(updates)) {
    set(sheetIds.simulador, ref, value);
  }
}

function init() {
  postMessage({ type: 'progress', payload: { message: 'Montando modelo financeiro...' } });
  hf = HyperFormula.buildFromSheets(buildSheets(), {
    licenseKey: 'gpl-v3',
    useArrayArithmetic: true,
  });
  sheetIds = {
    simulador: hf.getSheetId('SIMULADOR'),
    controle: hf.getSheetId('Controle'),
  };
  postMessage({ type: 'ready', payload: { results: currentResults() } });
}

self.addEventListener('message', (event) => {
  if (event.data.type !== 'simulate') {
    return;
  }
  try {
    applyUpdates(event.data.payload.updates);
    set(sheetIds.controle, 'M3', 15000);
    const seek = goalSeek(0.1);
    postMessage({
      type: 'result',
      payload: {
        results: currentResults(),
        iterations: seek.iterations,
      },
    });
  } catch (error) {
    postMessage({
      type: 'error',
      payload: { message: error.message || 'Erro ao simular.' },
    });
  }
});

init();
