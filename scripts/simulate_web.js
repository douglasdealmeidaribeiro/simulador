const fs = require('fs');
const path = require('path');
const { HyperFormula } = require('hyperformula');

function loadBrowserAssetObject(filePath, globalName) {
  global.window = {};
  const source = fs.readFileSync(filePath, 'utf8');
  eval(source);
  const value = global.window[globalName];
  if (!value) {
    throw new Error(`Asset ${globalName} nao encontrado em ${filePath}`);
  }
  return value;
}

function colToNum(col) {
  return [...col].reduce((acc, char) => acc * 26 + char.charCodeAt(0) - 64, 0);
}

function refToAddress(ref, sheet) {
  const match = ref.match(/^([A-Z]+)(\d+)$/);
  if (!match) {
    throw new Error(`Referencia invalida: ${ref}`);
  }
  return { sheet, row: Number(match[2]) - 1, col: colToNum(match[1]) - 1 };
}

function valueOf(value) {
  if (value && typeof value === 'object' && value.value) {
    throw new Error(value.message || value.value);
  }
  return value;
}

function buildSheets(workbookModel) {
  const sheets = {};
  for (const sheet of workbookModel.sheets) {
    const rows = Array.from({ length: sheet.maxRow }, () => []);
    for (const [ref, value] of Object.entries(sheet.cells)) {
      const address = refToAddress(ref, 0);
      rows[address.row][address.col] = value;
    }
    sheets[sheet.name] = rows;
  }
  return sheets;
}

const SHELTER_SENSITIVE_ITEMS = new Set([
  'DIFAL',
  'REINVESTIMENTOS  REPOSIÇÃO DE INFRAESTRUTURA 1 ANO',
  'REINVESTIMENTOS REPOSIÇÃO DE INFRAESTRUTURA 5 ANOS ',
  'REINVESTIMENTOS REPOSIÇÃO DE INFRAESTRUTURA 10 ANOS',
  'REINVESTIMENTOS REPOSIÇÃO DE INFRAESTRUTURA 15 ANOS',
]);

function staticValue(sheet, ref) {
  const value = sheet.cells[ref];
  if (typeof value === 'string' && value.startsWith('=')) {
    return sheet.cache?.[ref] ?? 0;
  }
  return value ?? sheet.cache?.[ref] ?? 0;
}

function buildIndirectBudgetModel(workbookModel) {
  const simulador = workbookModel.sheets.find((sheet) => sheet.name === 'SIMULADOR');
  const mensal = workbookModel.sheets.find((sheet) => sheet.name.endsWith('(Mensal)'));
  const origem = workbookModel.sheets.find((sheet) => sheet.name.endsWith('(Origem)'));

  const itemToBudgetRow = new Map();
  for (let row = 11; row <= 102; row += 1) {
    const item = staticValue(simulador, `E${row}`);
    if (item) {
      itemToBudgetRow.set(item, row);
    }
  }

  const centerToToggleRow = new Map();
  for (let row = 11; row <= 26; row += 1) {
    const center = staticValue(simulador, `B${row}`);
    if (center) {
      centerToToggleRow.set(center, row);
    }
  }

  const rows = [];
  for (let row = 7; row <= 87; row += 1) {
    const item = staticValue(mensal, `A${row}`);
    const budgetRow = itemToBudgetRow.get(item);
    const center = budgetRow ? staticValue(simulador, `F${budgetRow}`) : null;
    const toggleRow = centerToToggleRow.get(center);
    const value = staticValue(mensal, `S${row}`);
    if (!item || !budgetRow || !toggleRow || typeof value !== 'number' || value === 0) {
      continue;
    }
    rows.push({
      item,
      value,
      budgetRow,
      toggleRef: `C${toggleRow}`,
      isInvestment: staticValue(mensal, `B${row}`) === 'Investimentos',
      isCost: staticValue(mensal, `B${row}`) === 'Custos_despesas',
      isReplacement: staticValue(mensal, `C${row}`) === 'Reposição_de_infraestrutura',
      includeInvestments: staticValue(mensal, `I${row}`) === 'SIM',
      includePlans: staticValue(mensal, `L${row}`) === 'SIM',
      shelterSensitive: SHELTER_SENSITIVE_ITEMS.has(item),
    });
  }

  return {
    rows,
    difalBase: staticValue(origem, 'D32'),
    replacements: {
      D62: staticValue(origem, 'D62'),
      D66: staticValue(origem, 'D66'),
      D68: staticValue(origem, 'D68'),
      D70: staticValue(origem, 'D70'),
    },
  };
}

function defaultUpdates(simConfig) {
  const updates = {};
  for (const field of simConfig.general || []) {
    updates[field.cell] = field.value;
  }
  for (const item of simConfig.costCenters || []) {
    updates[item.cell] = item.value;
  }
  for (const item of simConfig.budget || []) {
    updates[item.cell] = item.value;
  }
  return updates;
}

function runSimulation(workbookModel, mergedUpdates, debugRefs = []) {
  const hf = HyperFormula.buildFromSheets(buildSheets(workbookModel), {
    licenseKey: 'gpl-v3',
    useArrayArithmetic: true,
  });
  const sheetIds = {
    simulador: hf.getSheetId('SIMULADOR'),
    controle: hf.getSheetId('Controle'),
    origem: hf.getSheetId('Orçamento (Origem)'),
  };
  const indirectBudgetModel = buildIndirectBudgetModel(workbookModel);

  function get(sheet, ref) {
    return valueOf(hf.getCellValue(refToAddress(ref, sheet)));
  }

  function set(sheet, ref, value) {
    hf.setCellContents(refToAddress(ref, sheet), [[value]]);
  }

  function evaluatePrice(price) {
    set(sheetIds.controle, 'M3', price);
    return Number(get(sheetIds.controle, 'M13'));
  }

  function numberFromSim(ref, fallback = 0) {
    const value = get(sheetIds.simulador, ref);
    return typeof value === 'number' && Number.isFinite(value) ? value : fallback;
  }

  function isSim(ref) {
    return String(get(sheetIds.simulador, ref)).trim().toUpperCase() === 'SIM';
  }

  function indirectRowValue(row) {
    if (!isSim(row.toggleRef)) {
      return 0;
    }
    if (row.shelterSensitive && !isSim('C15')) {
      return 0;
    }
    return row.value * (1 + numberFromSim(`G${row.budgetRow}`));
  }

  function indirectSum(predicate) {
    return indirectBudgetModel.rows
      .filter(predicate)
      .reduce((sum, row) => sum + indirectRowValue(row), 0);
  }

  function applyIndirectBudgets() {
    const riskInvestments = numberFromSim('C3');
    const riskOperations = numberFromSim('C4');
    const investmentSum = indirectSum((row) => row.isInvestment && row.includeInvestments);
    const costSum = indirectSum((row) => row.isCost && row.includeInvestments);
    const replacementSum = indirectSum((row) => row.isReplacement && row.includeInvestments);
    const plansSum = indirectSum((row) => row.includePlans);

    set(sheetIds.origem, 'D23', investmentSum * riskInvestments * (isSim('C25') ? 1 : 0) * (1 + numberFromSim('G31')));
    set(sheetIds.origem, 'D24', costSum * riskOperations * (isSim('C17') ? 1 : 0) * (1 + numberFromSim('G32')));
    set(sheetIds.origem, 'D25', replacementSum * riskInvestments * (isSim('C17') ? 1 : 0) * (1 + numberFromSim('G33')));
    set(sheetIds.origem, 'D32', indirectBudgetModel.difalBase * (isSim('C13') ? 1 : 0) * (isSim('C15') ? 1 : 0) * (1 + numberFromSim('G40')));
    set(sheetIds.origem, 'D59', plansSum * 0.05 * (isSim('C25') ? 1 : 0) * (1 + numberFromSim('G67')));

    for (const [ref, value] of Object.entries(indirectBudgetModel.replacements)) {
      const budgetRow = Number(ref.slice(1)) + 8;
      set(sheetIds.origem, ref, value * (isSim('C25') ? 1 : 0) * (isSim('C15') ? 1 : 0) * (1 + numberFromSim(`G${budgetRow}`)));
    }
  }

  function goalSeek(goal = 0.1) {
    let price = Number(get(sheetIds.controle, 'M3'));
    if (!Number.isFinite(price) || price < 0) {
      price = 15000;
    }

    let residual = evaluatePrice(price) - goal;
    let bestPrice = price;
    let bestResidual = residual;
    let step = Math.max(1, Math.abs(price) * 0.01);
    const tolerance = 0.0001;

    for (let i = 0; i < 100; i += 1) {
      residual = evaluatePrice(price) - goal;
      if (Math.abs(residual) < Math.abs(bestResidual)) {
        bestResidual = residual;
        bestPrice = price;
      }
      if (Math.abs(residual) <= tolerance) {
        break;
      }

      const probePrice = Math.max(0, price + step);
      const probeResidual = evaluatePrice(probePrice) - goal;
      const denominator = probeResidual - residual;

      let nextPrice;
      if (Number.isFinite(denominator) && Math.abs(denominator) > 1e-12) {
        nextPrice = price - residual * (probePrice - price) / denominator;
      } else {
        nextPrice = residual > 0 ? price - step : price + step;
      }

      if (!Number.isFinite(nextPrice)) {
        nextPrice = residual > 0 ? price - step : price + step;
      }

      nextPrice = Math.max(0, nextPrice);
      const maxJump = Math.max(500, Math.abs(price) * 0.5);
      const jump = nextPrice - price;
      if (Math.abs(jump) > maxJump) {
        nextPrice = price + Math.sign(jump) * maxJump;
      }

      step = Math.max(0.5, Math.abs(nextPrice - price) * 0.5);
      price = nextPrice;
    }

    set(sheetIds.controle, 'M3', bestPrice);
    return { price: bestPrice, residual: bestResidual };
  }

  for (const [ref, value] of Object.entries(mergedUpdates)) {
    set(sheetIds.simulador, ref, value);
  }
  applyIndirectBudgets();

  set(sheetIds.controle, 'M3', 15000);
  const seek = goalSeek(0.1);

  const result = {
    annual: Number(get(sheetIds.simulador, 'G3')),
    monthly: Number(get(sheetIds.simulador, 'G4')),
    target: Number(get(sheetIds.controle, 'M13')),
    price: Number(get(sheetIds.controle, 'M3')),
    residual: Number(seek.residual),
  };

  if (debugRefs.length > 0) {
    const debug = {};
    for (const ref of debugRefs) {
      const clean = String(ref || '').trim();
      if (!clean) {
        continue;
      }
      debug[clean] = Number(get(sheetIds.origem, clean));
    }
    result.debugOrigem = debug;
  }

  return result;
}

function main() {
  const root = path.resolve(__dirname, '..');
  const updatesPath = process.argv[2];
  const debugArg = process.argv[3] || '';
  if (!updatesPath) {
    throw new Error('Uso: node scripts/simulate_web.js <arquivo-json-de-updates> [ref1,ref2,...]');
  }

  const workbookModel = loadBrowserAssetObject(path.join(root, 'assets', 'workbook-model.js'), 'WORKBOOK_MODEL');
  const simConfig = loadBrowserAssetObject(path.join(root, 'assets', 'simulador-config.js'), 'SIMULADOR_CONFIG');
  const rawUpdates = fs.readFileSync(path.resolve(updatesPath), 'utf8').replace(/^\uFEFF/, '');
  const scenarioUpdates = JSON.parse(rawUpdates);

  const mergedUpdates = {
    ...defaultUpdates(simConfig),
    ...(scenarioUpdates || {}),
  };

  const debugRefs = debugArg
    .split(',')
    .map((item) => item.trim().toUpperCase())
    .filter(Boolean);

  const result = runSimulation(workbookModel, mergedUpdates, debugRefs);
  process.stdout.write(`${JSON.stringify(result)}\n`);
}

try {
  main();
} catch (error) {
  process.stderr.write(`${error.message || error}\n`);
  process.exit(1);
}
