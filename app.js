const config = window.SIMULADOR_CONFIG;

const state = new Map();
let engineReady = false;
let running = false;
let lastResults = null;

const engineStatus = document.querySelector('#engineStatus');
const generateButton = document.querySelector('#generateButton');
const annualResult = document.querySelector('#annualResult');
const monthlyResult = document.querySelector('#monthlyResult');
const targetResult = document.querySelector('#targetResult');
const priceResult = document.querySelector('#priceResult');
const worker = new Worker('worker.js?v=20260710-5');

generateButton.disabled = true;

const currencyFormatter = new Intl.NumberFormat('pt-BR', {
  style: 'currency',
  currency: 'BRL',
  maximumFractionDigits: 2,
});

const decimalFormatter = new Intl.NumberFormat('pt-BR', {
  maximumFractionDigits: 6,
});

function normalizeText(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase();
}

function formatPercentInput(value) {
  const number = Number(value || 0) * 100;
  return Number.isInteger(number) ? String(number) : number.toFixed(4).replace(/0+$/, '').replace(/\.$/, '');
}

function parsePercentInput(value) {
  const normalized = String(value).replace(',', '.').trim();
  const number = Number(normalized);
  return Number.isFinite(number) ? number / 100 : 0;
}

function setStatus(message, kind = '') {
  engineStatus.textContent = message;
  engineStatus.className = `status ${kind}`.trim();
}

function updateResults(results) {
  lastResults = results;
  annualResult.textContent = currencyFormatter.format(results.annual || 0);
  monthlyResult.textContent = currencyFormatter.format(results.monthly || 0);
  targetResult.textContent = decimalFormatter.format(results.target || 0);
  priceResult.textContent = currencyFormatter.format(results.price || 0);
}

function updateState(cell, value) {
  state.set(cell, value);
}

function createSelect(options, value, onChange) {
  const select = document.createElement('select');
  for (const optionValue of options) {
    const option = document.createElement('option');
    option.value = optionValue;
    option.textContent = optionValue;
    select.append(option);
  }
  select.value = value;
  select.addEventListener('change', () => onChange(select.value));
  return select;
}

function renderGeneral() {
  const host = document.querySelector('#generalFields');
  for (const field of config.general) {
    const row = document.createElement('div');
    row.className = 'field';
    const label = document.createElement('label');
    label.textContent = field.label;
    row.append(label);

    if (field.cell === 'C6') {
      const select = createSelect(['MEDIANA', 'MÍNIMO'], field.value, (value) => updateState(field.cell, value));
      row.append(select);
      updateState(field.cell, field.value);
    } else {
      const input = document.createElement('input');
      input.type = 'number';
      input.step = '0.01';
      input.value = formatPercentInput(field.value);
      input.addEventListener('input', () => updateState(field.cell, parsePercentInput(input.value)));
      row.append(input);
      updateState(field.cell, Number(field.value || 0));
    }

    host.append(row);
  }
}

function renderCostCenters() {
  const host = document.querySelector('#costCenters');
  for (const item of config.costCenters) {
    const row = document.createElement('label');
    row.className = 'toggle-row';
    const span = document.createElement('span');
    span.textContent = item.label;
    const select = createSelect(['SIM', 'NÃO'], item.value, (value) => updateState(item.cell, value));
    row.append(span, select);
    host.append(row);
    updateState(item.cell, item.value);
  }
}

function renderBudget() {
  const host = document.querySelector('#budgetRows');
  for (const item of config.budget) {
    const tr = document.createElement('tr');
    tr.dataset.search = normalizeText(`${item.item} ${item.center}`);

    const name = document.createElement('td');
    name.textContent = item.item;
    const center = document.createElement('td');
    center.textContent = item.center;
    const adjustment = document.createElement('td');
    const input = document.createElement('input');
    input.type = 'number';
    input.step = '0.01';
    input.value = formatPercentInput(item.value);
    input.addEventListener('input', () => updateState(item.cell, parsePercentInput(input.value)));
    adjustment.append(input);

    tr.append(name, center, adjustment);
    host.append(tr);
    updateState(item.cell, Number(item.value || 0));
  }
}

function applyFilter() {
  const query = normalizeText(document.querySelector('#budgetFilter').value);
  for (const row of document.querySelectorAll('#budgetRows tr')) {
    row.classList.toggle('hidden', query && !row.dataset.search.includes(query));
  }
}

function xmlEscape(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function cell(value) {
  if (typeof value === 'number' && Number.isFinite(value)) {
    return `<Cell><Data ss:Type="Number">${value}</Data></Cell>`;
  }
  return `<Cell><Data ss:Type="String">${xmlEscape(value)}</Data></Cell>`;
}

function row(values) {
  return `<Row>${values.map(cell).join('')}</Row>`;
}

function worksheet(name, rows) {
  return [
    `<Worksheet ss:Name="${xmlEscape(name.slice(0, 31))}">`,
    '<Table>',
    rows.map(row).join(''),
    '</Table>',
    '</Worksheet>',
  ].join('');
}

function currentStateValue(cellRef) {
  return state.has(cellRef) ? state.get(cellRef) : '';
}

function buildWorkbookXml(results) {
  const now = new Date();
  const resumo = [
    ['Simulador de Modelagem Econômico-Financeira'],
    ['Arquivo gerado no navegador', now.toLocaleString('pt-BR')],
    ['Origem do modelo', config.source],
    [],
    ['Indicador', 'Valor'],
    ['Contraprestação anual calculada', results.annual],
    ['Contraprestação mensal calculada', results.monthly],
    ['Preço usado no motor local', results.price],
    ['Meta/VPL de controle', results.target],
  ];

  const gerais = [
    ['Campo', 'Célula', 'Valor informado'],
    ...config.general.map((item) => [item.label, item.cell, currentStateValue(item.cell)]),
  ];

  const centros = [
    ['Centro de custos', 'Célula', 'Incluído'],
    ...config.costCenters.map((item) => [item.label, item.cell, currentStateValue(item.cell)]),
  ];

  const orcamento = [
    ['Orçamento de origem', 'Centro de custos', 'Célula', 'Ajuste manual'],
    ...config.budget.map((item) => [item.item, item.center, item.cell, currentStateValue(item.cell)]),
  ];

  return [
    '<?xml version="1.0" encoding="UTF-8"?>',
    '<?mso-application progid="Excel.Sheet"?>',
    '<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"',
    ' xmlns:o="urn:schemas-microsoft-com:office:office"',
    ' xmlns:x="urn:schemas-microsoft-com:office:excel"',
    ' xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"',
    ' xmlns:html="http://www.w3.org/TR/REC-html40">',
    worksheet('Resumo', resumo),
    worksheet('Geral', gerais),
    worksheet('Centros de custo', centros),
    worksheet('Ajustes orcamento', orcamento),
    '</Workbook>',
  ].join('');
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  document.body.append(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function downloadWorkbook(results) {
  const xml = buildWorkbookXml(results);
  const blob = new Blob([xml], { type: 'application/vnd.ms-excel;charset=utf-8' });
  const stamp = new Date().toISOString().slice(0, 10);
  downloadBlob(blob, `simulador-modelagem-calculado-${stamp}.xls`);
}

function setRunning(value) {
  running = value;
  generateButton.disabled = !engineReady || running;
}

function generateWorkbook() {
  if (!engineReady) {
    setStatus('Motor de cálculo ainda está carregando...', 'error');
    return;
  }

  setRunning(true);
  setStatus('Calculando no navegador...');
  worker.postMessage({
    type: 'simulate',
    payload: {
      updates: Object.fromEntries(state),
    },
  });
}

worker.addEventListener('message', (event) => {
  const { type, payload } = event.data;
  if (type === 'progress') {
    setStatus(payload.message);
    return;
  }
  if (type === 'ready') {
    engineReady = true;
    updateResults(payload.results);
    setRunning(false);
    setStatus('Motor local pronto. Não há backend ou servidor Excel.', 'ready');
    return;
  }
  if (type === 'result') {
    updateResults(payload.results);
    downloadWorkbook(payload.results);
    setRunning(false);
    setStatus('Planilha de resultados gerada no navegador.', 'ready');
    return;
  }
  if (type === 'error') {
    setRunning(false);
    setStatus(payload.message || 'Erro ao simular.', 'error');
  }
});

worker.addEventListener('error', () => {
  engineReady = false;
  setRunning(false);
  setStatus('Não foi possível carregar o motor local. Publique em um servidor estático ou abra via localhost.', 'error');
});

renderGeneral();
renderCostCenters();
renderBudget();
document.querySelector('#budgetFilter').addEventListener('input', applyFilter);
generateButton.addEventListener('click', generateWorkbook);
