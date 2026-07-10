const config = window.SIMULADOR_CONFIG;

const state = new Map();
let worker;
let engineReady = false;

const annualResult = document.querySelector('#annualResult');
const monthlyResult = document.querySelector('#monthlyResult');
const engineStatus = document.querySelector('#engineStatus');
const simulateButton = document.querySelector('#simulateButton');
const downloadButton = document.querySelector('#downloadButton');

let lastResults = config.results;
let isDirty = false;
let pendingDownload = false;

const currencyFormatter = new Intl.NumberFormat('pt-BR', {
  style: 'currency',
  currency: 'BRL',
  minimumFractionDigits: 2,
  maximumFractionDigits: 2,
});

function normalizeText(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase();
}

function formatMoney(value) {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    return 'R$ --';
  }
  return currencyFormatter.format(value);
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

function renderResults(results) {
  lastResults = results;
  annualResult.textContent = formatMoney(results.annual);
  monthlyResult.textContent = formatMoney(results.monthly);
}

function updateState(cell, value) {
  state.set(cell, value);
  isDirty = true;
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

function startWorker() {
  worker = new Worker('worker.js');

  worker.addEventListener('message', (event) => {
    const { type, payload } = event.data;

    if (type === 'ready') {
      engineReady = true;
      simulateButton.disabled = false;
      downloadButton.disabled = false;
      setStatus('Motor pronto', 'ready');
      renderResults(payload.results);
      isDirty = false;
    }

    if (type === 'progress') {
      setStatus(payload.message);
    }

    if (type === 'result') {
      simulateButton.disabled = false;
      downloadButton.disabled = false;
      setStatus(`Simulação concluída em ${payload.iterations} iterações`, 'ready');
      renderResults(payload.results);
      isDirty = false;
      if (pendingDownload) {
        pendingDownload = false;
        downloadSimulatedWorkbook();
      }
    }

    if (type === 'error') {
      simulateButton.disabled = false;
      downloadButton.disabled = false;
      pendingDownload = false;
      setStatus(payload.message, 'error');
    }
  });
}

function simulate(downloadAfter = false) {
  if (!engineReady) {
    return;
  }
  pendingDownload = downloadAfter;
  simulateButton.disabled = true;
  downloadButton.disabled = true;
  setStatus('Simulando...');
  worker.postMessage({
    type: 'simulate',
    payload: {
      updates: Object.fromEntries(state),
    },
  });
}

function parseXml(xml) {
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  const error = doc.querySelector('parsererror');
  if (error) {
    throw new Error('Não foi possível ler a estrutura XML da planilha.');
  }
  return doc;
}

function serializeXml(doc) {
  return new XMLSerializer().serializeToString(doc);
}

function columnNumber(ref) {
  const match = ref.match(/^([A-Z]+)/);
  return [...match[1]].reduce((acc, char) => acc * 26 + char.charCodeAt(0) - 64, 0);
}

function rowNumber(ref) {
  return Number(ref.match(/\d+$/)[0]);
}

function ensureRow(doc, rowIndex) {
  const rows = [...doc.getElementsByTagName('row')];
  const existing = rows.find((row) => Number(row.getAttribute('r')) === rowIndex);
  if (existing) {
    return existing;
  }

  const sheetData = doc.getElementsByTagName('sheetData')[0];
  const row = doc.createElementNS('http://schemas.openxmlformats.org/spreadsheetml/2006/main', 'row');
  row.setAttribute('r', String(rowIndex));

  const nextRow = rows.find((candidate) => Number(candidate.getAttribute('r')) > rowIndex);
  sheetData.insertBefore(row, nextRow || null);
  return row;
}

function ensureCell(doc, ref) {
  const cells = [...doc.getElementsByTagName('c')];
  const existing = cells.find((cell) => cell.getAttribute('r') === ref);
  if (existing) {
    return existing;
  }

  const row = ensureRow(doc, rowNumber(ref));
  const cell = doc.createElementNS('http://schemas.openxmlformats.org/spreadsheetml/2006/main', 'c');
  cell.setAttribute('r', ref);

  const targetCol = columnNumber(ref);
  const rowCells = [...row.getElementsByTagName('c')];
  const nextCell = rowCells.find((candidate) => columnNumber(candidate.getAttribute('r')) > targetCol);
  row.insertBefore(cell, nextCell || null);
  return cell;
}

function removeChildren(cell, names) {
  for (const name of names) {
    for (const child of [...cell.getElementsByTagName(name)]) {
      if (child.parentNode === cell) {
        cell.removeChild(child);
      }
    }
  }
}

function setCell(doc, ref, value, options = {}) {
  const cell = ensureCell(doc, ref);
  const preserveFormula = Boolean(options.preserveFormula);
  const hasFormula = [...cell.getElementsByTagName('f')].some((child) => child.parentNode === cell);

  if (!preserveFormula || !hasFormula) {
    removeChildren(cell, ['f', 'v', 'is']);
  } else {
    removeChildren(cell, ['v']);
  }

  if (typeof value === 'string') {
    cell.setAttribute('t', 'inlineStr');
    const inline = doc.createElementNS('http://schemas.openxmlformats.org/spreadsheetml/2006/main', 'is');
    const text = doc.createElementNS('http://schemas.openxmlformats.org/spreadsheetml/2006/main', 't');
    text.textContent = value;
    inline.append(text);
    cell.append(inline);
    return;
  }

  cell.removeAttribute('t');
  const number = doc.createElementNS('http://schemas.openxmlformats.org/spreadsheetml/2006/main', 'v');
  number.textContent = String(Number(value || 0));
  cell.append(number);
}

function normalizeWorkbookTarget(target) {
  return `xl/${target}`.replace(/\/+/g, '/');
}

async function workbookSheetPaths(zip) {
  const workbookXml = await zip.file('xl/workbook.xml').async('string');
  const relsXml = await zip.file('xl/_rels/workbook.xml.rels').async('string');
  const workbookDoc = parseXml(workbookXml);
  const relsDoc = parseXml(relsXml);

  const rels = new Map(
    [...relsDoc.getElementsByTagName('Relationship')].map((rel) => [
      rel.getAttribute('Id'),
      normalizeWorkbookTarget(rel.getAttribute('Target')),
    ]),
  );

  const paths = new Map();
  for (const sheet of workbookDoc.getElementsByTagName('sheet')) {
    const relId = sheet.getAttribute('r:id') || sheet.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id');
    paths.set(sheet.getAttribute('name'), rels.get(relId));
  }
  return { workbookDoc, paths };
}

async function updateWorksheet(zip, path, updates) {
  const file = zip.file(path);
  if (!file) {
    throw new Error(`A aba ${path} não foi encontrada na planilha.`);
  }
  const doc = parseXml(await file.async('string'));
  for (const update of updates) {
    setCell(doc, update.ref, update.value, update.options);
  }
  zip.file(path, serializeXml(doc));
}

function setFullRecalculation(workbookDoc) {
  const workbook = workbookDoc.documentElement;
  let calcPr = workbookDoc.getElementsByTagName('calcPr')[0];
  if (!calcPr) {
    calcPr = workbookDoc.createElementNS('http://schemas.openxmlformats.org/spreadsheetml/2006/main', 'calcPr');
    workbook.append(calcPr);
  }
  calcPr.setAttribute('calcMode', 'auto');
  calcPr.setAttribute('fullCalcOnLoad', '1');
  calcPr.setAttribute('forceFullCalc', '1');
}

function workbookInputs() {
  return [...state.entries()].map(([ref, value]) => ({ ref, value }));
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

async function downloadSimulatedWorkbook() {
  try {
    simulateButton.disabled = true;
    downloadButton.disabled = true;
    setStatus('Gerando planilha...');

    const response = await fetch(config.source);
    if (!response.ok) {
      throw new Error('Não foi possível carregar a planilha original.');
    }

    const zip = await JSZip.loadAsync(await response.arrayBuffer());
    const { workbookDoc, paths } = await workbookSheetPaths(zip);
    const simuladorPath = paths.get('SIMULADOR');
    const controlePath = paths.get('Controle');

    await updateWorksheet(zip, simuladorPath, [
      ...workbookInputs(),
      { ref: 'G3', value: lastResults.annual, options: { preserveFormula: true } },
      { ref: 'G4', value: lastResults.monthly, options: { preserveFormula: true } },
    ]);
    await updateWorksheet(zip, controlePath, [
      { ref: 'M3', value: lastResults.annual },
      { ref: 'M13', value: 0.1, options: { preserveFormula: true } },
    ]);

    setFullRecalculation(workbookDoc);
    zip.file('xl/workbook.xml', serializeXml(workbookDoc));

    const blob = await zip.generateAsync({
      type: 'blob',
      mimeType: 'application/vnd.ms-excel.sheet.macroEnabled.12',
      compression: 'DEFLATE',
    });
    downloadBlob(blob, 'simulador-modelagem-simulado.xlsm');
    setStatus('Planilha simulada gerada', 'ready');
  } catch (error) {
    setStatus(error.message || 'Erro ao gerar a planilha.', 'error');
  } finally {
    simulateButton.disabled = false;
    downloadButton.disabled = false;
  }
}

function requestWorkbookDownload() {
  if (!engineReady) {
    return;
  }
  if (isDirty) {
    simulate(true);
    return;
  }
  downloadSimulatedWorkbook();
}

renderResults(config.results);
renderGeneral();
renderCostCenters();
renderBudget();
isDirty = false;
document.querySelector('#budgetFilter').addEventListener('input', applyFilter);
simulateButton.addEventListener('click', () => simulate());
downloadButton.addEventListener('click', requestWorkbookDownload);
startWorker();
