const http = require('http');
const fs = require('fs');
const os = require('os');
const path = require('path');
const { spawn } = require('child_process');

const ROOT = path.resolve(__dirname, '..');
const PORT = Number(process.env.PORT || 3000);
const DEFAULT_WORKBOOK = fs.readdirSync(ROOT).find((file) => file.toLowerCase().endsWith('.xlsm'));
const WORKBOOK_PATH = process.env.WORKBOOK_PATH
  ? path.resolve(process.env.WORKBOOK_PATH)
  : (DEFAULT_WORKBOOK ? path.join(ROOT, DEFAULT_WORKBOOK) : '');
const FRONTEND_ORIGIN = process.env.FRONTEND_ORIGIN || '*';
const MAX_BODY_BYTES = 1024 * 1024;
const EXCEL_TIMEOUT_MS = Number(process.env.EXCEL_TIMEOUT_MS || 5 * 60 * 1000);

function sendCors(res) {
  res.setHeader('Access-Control-Allow-Origin', FRONTEND_ORIGIN);
  res.setHeader('Access-Control-Allow-Methods', 'GET,POST,OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
}

function json(res, status, payload) {
  sendCors(res);
  res.writeHead(status, { 'Content-Type': 'application/json; charset=utf-8' });
  res.end(JSON.stringify(payload));
}

function readBody(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    let size = 0;
    req.on('data', (chunk) => {
      size += chunk.length;
      if (size > MAX_BODY_BYTES) {
        reject(new Error('Payload muito grande.'));
        req.destroy();
        return;
      }
      chunks.push(chunk);
    });
    req.on('end', () => resolve(Buffer.concat(chunks).toString('utf8')));
    req.on('error', reject);
  });
}

function safeUnlink(filePath) {
  try {
    if (filePath && fs.existsSync(filePath)) {
      fs.unlinkSync(filePath);
    }
  } catch (_error) {
    // Best effort cleanup.
  }
}

function runExcelSimulation(inputPath, outputPath) {
  return new Promise((resolve, reject) => {
    const scriptPath = path.join(__dirname, 'simulate-excel.ps1');
    const child = spawn('powershell.exe', [
      '-NoProfile',
      '-ExecutionPolicy',
      'Bypass',
      '-File',
      scriptPath,
      '-WorkbookPath',
      WORKBOOK_PATH,
      '-InputJsonPath',
      inputPath,
      '-OutputWorkbookPath',
      outputPath,
    ], { windowsHide: false, stdio: ['ignore', 'pipe', 'pipe'] });

    let stdout = '';
    let stderr = '';
    const timeout = setTimeout(() => {
      child.kill();
      reject(new Error('Tempo limite excedido ao calcular no Excel.'));
    }, EXCEL_TIMEOUT_MS);

    child.stdout.on('data', (chunk) => { stdout += chunk.toString(); });
    child.stderr.on('data', (chunk) => { stderr += chunk.toString(); });
    child.on('error', reject);
    child.on('close', (code) => {
      clearTimeout(timeout);
      if (code !== 0) {
        reject(new Error(stderr || stdout || `Excel finalizou com código ${code}.`));
        return;
      }
      resolve(stdout);
    });
  });
}

async function handleSimulation(req, res) {
  const id = `${Date.now()}-${Math.random().toString(16).slice(2)}`;
  const workDir = path.join(os.tmpdir(), 'simulador-excel');
  fs.mkdirSync(workDir, { recursive: true });
  const inputPath = path.join(workDir, `${id}.json`);
  const outputPath = path.join(workDir, `${id}.xlsm`);

  try {
    if (!fs.existsSync(WORKBOOK_PATH)) {
      throw new Error(`Planilha base não encontrada: ${WORKBOOK_PATH}`);
    }

    const payload = JSON.parse(await readBody(req));
    if (!payload || typeof payload.updates !== 'object' || Array.isArray(payload.updates)) {
      throw new Error('Envie um JSON no formato { "updates": { "C3": 0.0197 } }.');
    }

    fs.writeFileSync(inputPath, JSON.stringify(payload), 'utf8');
    await runExcelSimulation(inputPath, outputPath);

    const file = fs.readFileSync(outputPath);
    sendCors(res);
    res.writeHead(200, {
      'Content-Type': 'application/vnd.ms-excel.sheet.macroEnabled.12',
      'Content-Disposition': 'attachment; filename="simulador-modelagem-calculado.xlsm"',
      'Content-Length': file.length,
    });
    res.end(file);
  } catch (error) {
    json(res, 500, { error: error.message || 'Erro ao calcular no Excel.' });
  } finally {
    safeUnlink(inputPath);
    safeUnlink(outputPath);
  }
}

const server = http.createServer(async (req, res) => {
  if (req.method === 'OPTIONS') {
    sendCors(res);
    res.writeHead(204);
    res.end();
    return;
  }

  if (req.method === 'GET' && req.url === '/health') {
    json(res, 200, { ok: true });
    return;
  }

  if (req.method === 'POST' && req.url === '/api/simular') {
    await handleSimulation(req, res);
    return;
  }

  json(res, 404, { error: 'Rota não encontrada.' });
});

server.listen(PORT, () => {
  console.log(`Simulador Excel API em http://localhost:${PORT}`);
});
