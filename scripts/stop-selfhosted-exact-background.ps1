param(
  [string]$TunnelName = 'simulador-api'
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$runtimeDir = Join-Path $root '.runtime'
$backendPidFile = Join-Path $runtimeDir 'backend.pid'
$tunnelPidFile = Join-Path $runtimeDir 'tunnel.pid'

function Stop-FromPidFile {
  param([string]$Path, [string]$Name)

  if (-not (Test-Path -LiteralPath $Path)) {
    Write-Host "${Name}: nao estava em execucao"
    return
  }

  $pidRaw = Get-Content -LiteralPath $Path -Raw
  if (-not [int]::TryParse($pidRaw, [ref]$null)) {
    Remove-Item -LiteralPath $Path -Force -ErrorAction SilentlyContinue
    Write-Host "${Name}: pid invalido removido"
    return
  }

  $procId = [int]$pidRaw
  $proc = Get-Process -Id $procId -ErrorAction SilentlyContinue
  if ($null -ne $proc) {
    Stop-Process -Id $procId -Force -ErrorAction SilentlyContinue
    Write-Host "${Name} parado (PID $procId)"
  } else {
    Write-Host "${Name}: processo nao encontrado"
  }
  Remove-Item -LiteralPath $Path -Force -ErrorAction SilentlyContinue
}

function Stop-ResidualProcesses {
  param([string]$Tunnel)

  $residual = Get-CimInstance Win32_Process | Where-Object {
    ($_.Name -ieq 'node.exe' -and $_.CommandLine -match 'backend/server.js') -or
    ($_.Name -ieq 'cloudflared.exe' -and $_.CommandLine -match "tunnel run\s+$Tunnel")
  }

  foreach ($p in $residual) {
    Stop-Process -Id $p.ProcessId -Force -ErrorAction SilentlyContinue
    Write-Host "Residual parado ($($p.Name) PID $($p.ProcessId))"
  }
}

Stop-FromPidFile -Path $tunnelPidFile -Name 'Tunnel'
Stop-FromPidFile -Path $backendPidFile -Name 'Backend'
Stop-ResidualProcesses -Tunnel $TunnelName
