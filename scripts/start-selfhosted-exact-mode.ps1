param(
  [string]$TunnelName = 'simulador-api',
  [string]$PublicHostname = '',
  [int]$Port = 3000,
  [string]$FrontendOrigin = '*'
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$backendScript = Join-Path $root 'backend\server.js'

function Resolve-CloudflaredPath {
  $command = Get-Command cloudflared -ErrorAction SilentlyContinue
  if ($command) {
    return $command.Source
  }

  $candidates = @(
    'C:\\Program Files\\cloudflared\\cloudflared.exe',
    'C:\\Program Files (x86)\\cloudflared\\cloudflared.exe'
  )

  foreach ($candidate in $candidates) {
    if (Test-Path -LiteralPath $candidate) {
      return $candidate
    }
  }

  throw 'cloudflared nao encontrado no PATH nem nas pastas padrao. Instale antes de continuar.'
}

if (-not (Test-Path -LiteralPath $backendScript)) {
  throw "Arquivo nao encontrado: $backendScript"
}

if (-not (Get-Command node -ErrorAction SilentlyContinue)) {
  throw 'Node.js nao encontrado no PATH.'
}

$cloudflaredPath = Resolve-CloudflaredPath

$backendEnv = "`$env:PORT=$Port; `$env:FRONTEND_ORIGIN='$FrontendOrigin'; node backend/server.js"
$backendArgs = @(
  '-NoExit'
  '-Command'
  $backendEnv
)

Write-Host "Iniciando API local em http://127.0.0.1:$Port ..."
$backendProcess = Start-Process -FilePath 'powershell.exe' -ArgumentList $backendArgs -WorkingDirectory $root -PassThru

Start-Sleep -Seconds 3

try {
  $health = Invoke-WebRequest -Uri "http://127.0.0.1:$Port/health" -UseBasicParsing -TimeoutSec 10
  if ($health.StatusCode -ne 200) {
    throw "Health check inesperado: $($health.StatusCode)"
  }
} catch {
  try { Stop-Process -Id $backendProcess.Id -Force } catch {}
  throw "A API local nao subiu corretamente: $($_.Exception.Message)"
}

Write-Host 'API local iniciada com sucesso.'
if ($PublicHostname) {
  Write-Host "Hostname publico esperado: https://$PublicHostname"
}
Write-Host "Iniciando Cloudflare Tunnel '$TunnelName' ..."
Write-Host 'Pressione Ctrl+C para encerrar o tunnel. O backend local sera finalizado junto.'

try {
  & $cloudflaredPath tunnel run $TunnelName
} finally {
  if ($backendProcess -and -not $backendProcess.HasExited) {
    try { Stop-Process -Id $backendProcess.Id -Force } catch {}
  }
}
