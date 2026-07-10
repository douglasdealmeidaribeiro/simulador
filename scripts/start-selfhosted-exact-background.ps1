param(
  [string]$TunnelName = 'simulador-api',
  [int]$Port = 3000,
  [string]$FrontendOrigin = '*'
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$runtimeDir = Join-Path $root '.runtime'
$backendPidFile = Join-Path $runtimeDir 'backend.pid'
$tunnelPidFile = Join-Path $runtimeDir 'tunnel.pid'

function Resolve-CloudflaredPath {
  $command = Get-Command cloudflared -ErrorAction SilentlyContinue
  if ($command) {
    return $command.Source
  }

  $candidates = @(
    'C:\Program Files\cloudflared\cloudflared.exe',
    'C:\Program Files (x86)\cloudflared\cloudflared.exe'
  )

  foreach ($candidate in $candidates) {
    if (Test-Path -LiteralPath $candidate) {
      return $candidate
    }
  }

  throw 'cloudflared nao encontrado no PATH nem nas pastas padrao.'
}

function Get-LiveProcessFromPidFile {
  param([string]$Path)
  if (-not (Test-Path -LiteralPath $Path)) {
    return $null
  }
  $pidRaw = Get-Content -LiteralPath $Path -Raw
  if (-not [int]::TryParse($pidRaw, [ref]$null)) {
    Remove-Item -LiteralPath $Path -Force -ErrorAction SilentlyContinue
    return $null
  }
  $pid = [int]$pidRaw
  $process = Get-Process -Id $pid -ErrorAction SilentlyContinue
  if ($null -eq $process) {
    Remove-Item -LiteralPath $Path -Force -ErrorAction SilentlyContinue
    return $null
  }
  return $process
}

if (-not (Get-Command node -ErrorAction SilentlyContinue)) {
  throw 'Node.js nao encontrado no PATH.'
}

$cloudflaredPath = Resolve-CloudflaredPath

if (-not (Test-Path -LiteralPath $runtimeDir)) {
  New-Item -ItemType Directory -Path $runtimeDir | Out-Null
}

$backendExisting = Get-LiveProcessFromPidFile -Path $backendPidFile
if ($null -eq $backendExisting) {
  $backendCommand = "$env:PORT=$Port; $env:FRONTEND_ORIGIN='$FrontendOrigin'; Set-Location '$root'; node backend/server.js"
  $backendProc = Start-Process -FilePath 'powershell.exe' -ArgumentList @('-NoProfile', '-WindowStyle', 'Hidden', '-ExecutionPolicy', 'Bypass', '-Command', $backendCommand) -WindowStyle Hidden -PassThru
  Set-Content -LiteralPath $backendPidFile -Value $backendProc.Id -Encoding ASCII
} else {
  $backendProc = $backendExisting
}

$ok = $false
for ($i = 0; $i -lt 20; $i++) {
  try {
    $resp = Invoke-WebRequest -Uri "http://127.0.0.1:$Port/health" -UseBasicParsing -TimeoutSec 5
    if ($resp.StatusCode -eq 200) {
      $ok = $true
      break
    }
  } catch {
    Start-Sleep -Milliseconds 500
  }
}
if (-not $ok) {
  throw 'Backend local nao respondeu no health check.'
}

$tunnelExisting = Get-LiveProcessFromPidFile -Path $tunnelPidFile
if ($null -eq $tunnelExisting) {
  $tunnelProc = Start-Process -FilePath $cloudflaredPath -ArgumentList @('tunnel', 'run', $TunnelName) -WindowStyle Hidden -PassThru
  Set-Content -LiteralPath $tunnelPidFile -Value $tunnelProc.Id -Encoding ASCII
} else {
  $tunnelProc = $tunnelExisting
}

Write-Host "Backend PID: $($backendProc.Id)"
Write-Host "Tunnel PID: $($tunnelProc.Id)"
Write-Host 'Modo self-hosted exato ativo em segundo plano.'
