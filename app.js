const config = window.SIMULADOR_CONFIG;

const state = new Map();
let worker;
let engineReady = false;

const annualResult = document.querySelector('#annualResult');
const monthlyResult = document.querySelector('#monthlyResult');
const engineStatus = document.querySelector('#engineStatus');
const simulateButton = document.querySelector('#simulateButton');

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
  annualResult.textContent = formatMoney(results.annual);
  monthlyResult.textContent = formatMoney(results.monthly);
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

function startWorker() {
  worker = new Worker('worker.js');

  worker.addEventListener('message', (event) => {
    const { type, payload } = event.data;

    if (type === 'ready') {
      engineReady = true;
      simulateButton.disabled = false;
      setStatus('Motor pronto', 'ready');
      renderResults(payload.results);
    }

    if (type === 'progress') {
      setStatus(payload.message);
    }

    if (type === 'result') {
      simulateButton.disabled = false;
      setStatus(`Simulação concluída em ${payload.iterations} iterações`, 'ready');
      renderResults(payload.results);
    }

    if (type === 'error') {
      simulateButton.disabled = false;
      setStatus(payload.message, 'error');
    }
  });
}

function simulate() {
  if (!engineReady) {
    return;
  }
  simulateButton.disabled = true;
  setStatus('Simulando...');
  worker.postMessage({
    type: 'simulate',
    payload: {
      updates: Object.fromEntries(state),
    },
  });
}

renderResults(config.results);
renderGeneral();
renderCostCenters();
renderBudget();
document.querySelector('#budgetFilter').addEventListener('input', applyFilter);
simulateButton.addEventListener('click', simulate);
startWorker();
