param(
  [string]$TunnelName = 'simulador-api',
  [int]$Port = 3000,
  [string]$FrontendOrigin = '*'
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$runtimeDir = Join-Path $root '.runtime'
$startScript = Join-Path $root 'scripts\start-selfhosted-exact-background.ps1'
$tunnelPidFile = Join-Path $runtimeDir 'tunnel.pid'
$lockPath = Join-Path $runtimeDir 'watchdog.lock'
$logPath = Join-Path $runtimeDir 'watchdog.log'

if (-not (Test-Path -LiteralPath $runtimeDir)) {
  New-Item -ItemType Directory -Path $runtimeDir | Out-Null
}

function Write-Log {
  param([string]$Message)
  $line = "$(Get-Date -Format s) $Message"
  Add-Content -LiteralPath $logPath -Value $line -Encoding ASCII
}

function Test-LocalHealth {
  try {
    $resp = Invoke-WebRequest -UseBasicParsing -Uri "http://127.0.0.1:$Port/health" -TimeoutSec 3
    return $resp.StatusCode -eq 200
  } catch {
    return $false
  }
}

function Test-TunnelProcess {
  if (Test-Path -LiteralPath $tunnelPidFile) {
    $pidRaw = Get-Content -LiteralPath $tunnelPidFile -Raw
    $procId = 0
    if ([int]::TryParse($pidRaw, [ref]$procId)) {
      $p = Get-Process -Id $procId -ErrorAction SilentlyContinue
      if ($null -ne $p) {
        return $true
      }
    }
  }

  # Fallback when process command line is not accessible by current user.
  $anyCloudflared = Get-Process cloudflared -ErrorAction SilentlyContinue | Select-Object -First 1
  return $null -ne $anyCloudflared
}

$lockHandle = $null
try {
  $lockHandle = [System.IO.File]::Open($lockPath, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
} catch {
  Write-Log 'watchdog em execucao, ignorando ciclo'
  exit 0
}

try {
  if (-not (Test-Path -LiteralPath $startScript)) {
    Write-Log "script de start nao encontrado: $startScript"
    exit 1
  }

  $localOk = Test-LocalHealth
  $tunnelOk = Test-TunnelProcess

  if ($localOk -and $tunnelOk) {
    Write-Log 'saude ok (backend+tunnel)'
    exit 0
  }

  Write-Log "saude degradada (localOk=$localOk, tunnelOk=$tunnelOk), tentando recuperar"
  & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $startScript -TunnelName $TunnelName -Port $Port -FrontendOrigin $FrontendOrigin

  if ($LASTEXITCODE -ne 0) {
    Write-Log "recuperacao falhou com codigo $LASTEXITCODE"
    exit $LASTEXITCODE
  }

  $localAfter = Test-LocalHealth
  $tunnelAfter = Test-TunnelProcess
  if ($localAfter -and $tunnelAfter) {
    Write-Log 'recuperacao concluida com sucesso'
    exit 0
  }

  Write-Log 'recuperacao executada, mas saude ainda degradada'
  exit 1
} finally {
  if ($lockHandle) {
    $lockHandle.Dispose()
  }
  Remove-Item -LiteralPath $lockPath -Force -ErrorAction SilentlyContinue
}
