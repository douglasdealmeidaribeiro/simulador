const config = window.SIMULADOR_CONFIG;
const apiBaseUrl = (window.SIMULADOR_API_URL || 'http://127.0.0.1:3000').replace(/\/$/, '');

const state = new Map();
let apiReady = false;

const engineStatus = document.querySelector('#engineStatus');
const generateButton = document.querySelector('#generateButton');
generateButton.disabled = true;

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

function apiUnavailableMessage() {
  const isLocalApi = /^https?:\/\/(localhost|127\.0\.0\.1)(:\d+)?$/i.test(apiBaseUrl);
  if (isLocalApi) {
    return 'Backend do Excel indisponível. Inicie a API em http://127.0.0.1:3000 ou configure uma URL pública HTTPS.';
  }
  return `Backend do Excel indisponível em ${apiBaseUrl}. Verifique se a API está online e liberada para este site.`;
}

async function checkApiHealth() {
  setStatus('Verificando backend do Excel...');
  try {
    const response = await fetch(`${apiBaseUrl}/health`, {
      cache: 'no-store',
      headers: { Accept: 'application/json' },
    });

    if (!response.ok) {
      throw new Error(apiUnavailableMessage());
    }

    apiReady = true;
    generateButton.disabled = false;
    generateButton.title = '';
    setStatus('Backend do Excel conectado', 'ready');
  } catch (_error) {
    apiReady = false;
    generateButton.disabled = true;
    generateButton.title = apiUnavailableMessage();
    setStatus(apiUnavailableMessage(), 'error');
  }
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

function filenameFromResponse(response) {
  const disposition = response.headers.get('content-disposition') || '';
  const match = disposition.match(/filename="?([^"]+)"?/i);
  return match ? match[1] : 'simulador-modelagem-calculado.xlsm';
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

async function generateWorkbook() {
  if (!apiReady) {
    setStatus(apiUnavailableMessage(), 'error');
    return;
  }

  generateButton.disabled = true;
  setStatus('Enviando para o Excel no servidor...');

  try {
    const response = await fetch(`${apiBaseUrl}/api/simular`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ updates: Object.fromEntries(state) }),
    });

    if (!response.ok) {
      let message = 'Não foi possível gerar a planilha calculada.';
      try {
        const payload = await response.json();
        message = payload.error || message;
      } catch (_error) {
        message = await response.text() || message;
      }
      throw new Error(message);
    }

    const blob = await response.blob();
    downloadBlob(blob, filenameFromResponse(response));
    setStatus('Planilha calculada pelo Excel gerada', 'ready');
  } catch (error) {
    const message = error instanceof TypeError && /fetch/i.test(error.message)
      ? apiUnavailableMessage()
      : (error.message || 'Erro ao gerar a planilha.');
    apiReady = false;
    setStatus(message, 'error');
  } finally {
    generateButton.disabled = !apiReady;
  }
}

renderGeneral();
renderCostCenters();
renderBudget();
checkApiHealth();
document.querySelector('#budgetFilter').addEventListener('input', applyFilter);
generateButton.addEventListener('click', generateWorkbook);
