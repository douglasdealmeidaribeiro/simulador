const config = window.SIMULADOR_CONFIG;
const apiConfig = window.API_CONFIG || {};

const state = new Map();
let engineReady = false;
let running = false;
let lastResults = null;

const engineStatus = document.querySelector('#engineStatus');
const generateButton = document.querySelector('#generateButton');
const annualResult = document.querySelector('#annualResult');
const monthlyResult = document.querySelector('#monthlyResult');
const backendBaseUrl = String(apiConfig.baseUrl || '').replace(/\/+$/, '');

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

function apiUrl(path) {
  if (!backendBaseUrl) {
    return path;
  }
  return `${backendBaseUrl}${path}`;
}

function apiHeaders(headers = {}) {
  const requestHeaders = { ...headers };
  if (backendBaseUrl.includes('ngrok')) {
    requestHeaders['ngrok-skip-browser-warning'] = 'true';
  }
  return requestHeaders;
}

function filenameFromContentDisposition(disposition, fallback) {
  const value = String(disposition || '');
  const match = value.match(/filename\*?=(?:UTF-8''|"?)([^";]+)/i);
  if (!match) {
    return fallback;
  }
  try {
    return decodeURIComponent(match[1].replace(/"/g, '').trim());
  } catch (_error) {
    return match[1].replace(/"/g, '').trim() || fallback;
  }
}

function parseBackendSummary(encodedHeader) {
  if (!encodedHeader) {
    return null;
  }
  try {
    return JSON.parse(decodeURIComponent(encodedHeader));
  } catch (_error) {
    return null;
  }
}

async function generateWithBackend(updates) {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 2 * 60 * 1000);

  try {
    const response = await fetch(apiUrl('/api/simular'), {
      method: 'POST',
      headers: apiHeaders({ 'Content-Type': 'application/json' }),
      body: JSON.stringify({ updates }),
      signal: controller.signal,
    });

    if (!response.ok) {
      let details = '';
      try {
        const payload = await response.json();
        details = payload?.error || '';
      } catch (_error) {
        details = '';
      }
      throw new Error(details || `Falha no backend Excel (${response.status}).`);
    }

    const blob = await response.blob();
    const contentDisposition = response.headers.get('Content-Disposition');
    const filename = filenameFromContentDisposition(contentDisposition, `orcamento-mensal-simulado-${new Date().toISOString().slice(0, 10)}.xlsx`);
    const summary = parseBackendSummary(response.headers.get('X-Simulation-Summary'));

    if (summary && typeof summary === 'object') {
      updateResults({
        annual: Number(summary.annual || 0),
        monthly: Number(summary.monthly || 0),
        target: Number(summary.target || 0),
        price: Number(summary.price || summary.annual || 0),
      });
    }

    downloadBlob(blob, filename);
    setStatus('Planilha Orçamento (Mensal) gerada no Excel com dados pós-simulação.', 'ready');
  } finally {
    clearTimeout(timeout);
  }
}

async function checkBackendHealth() {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 15000);

  try {
    const response = await fetch(apiUrl('/health'), {
      method: 'GET',
      headers: apiHeaders(),
      cache: 'no-store',
      signal: controller.signal,
    });
    if (!response.ok) {
      throw new Error(`Health check falhou (${response.status}).`);
    }
    const payload = await response.json();
    if (!payload?.ok) {
      throw new Error('Backend retornou resposta invalida.');
    }
  } finally {
    clearTimeout(timeout);
  }
}

function updateResults(results) {
  lastResults = results;
  annualResult.textContent = currencyFormatter.format(results.annual || 0);
  monthlyResult.textContent = currencyFormatter.format(results.monthly || 0);
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
    ['Preço usado na simulação', results.price],
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

function setRunning(value) {
  running = value;
  generateButton.disabled = !engineReady || running;
}

async function generateWorkbook() {
  if (!engineReady) {
    setStatus('Backend Excel indisponível. Inicie a API exata para simular.', 'error');
    return;
  }

  setRunning(true);
  const updates = Object.fromEntries(state);

  setStatus('Gerando Orçamento (Mensal) no Excel...');
  try {
    await generateWithBackend(updates);
  } catch (error) {
    setStatus(`Falha ao gerar planilha exata (${error.message || 'erro desconhecido'}).`, 'error');
  } finally {
    setRunning(false);
  }
}

async function initBackendMode() {
  setStatus('Verificando backend Excel...');
  try {
    await checkBackendHealth();
    engineReady = true;
    setRunning(false);
    setStatus('Backend Excel pronto. A simulação usa somente o resultado exato.', 'ready');
  } catch (error) {
    engineReady = false;
    setRunning(false);
    const detail = error?.message || 'backend indisponível';
    if (!backendBaseUrl) {
      setStatus(`Backend não encontrado (${detail}). Configure a API pública para usar a simulação exata.`, 'error');
      return;
    }
    setStatus(`Backend não encontrado em ${backendBaseUrl} (${detail}).`, 'error');
  }
}

renderGeneral();
renderCostCenters();
renderBudget();
document.querySelector('#budgetFilter').addEventListener('input', applyFilter);
generateButton.addEventListener('click', generateWorkbook);
initBackendMode();
