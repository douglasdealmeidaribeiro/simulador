param(
  [string]$TaskName = 'SimuladorExactSelfHostedMonitor',
  [int]$IntervalMinutes = 2,
  [string]$TunnelName = 'simulador-api',
  [int]$Port = 3000,
  [string]$FrontendOrigin = '*'
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$watchScript = Join-Path $root 'scripts\watch-selfhosted-health.ps1'

if (-not (Test-Path -LiteralPath $watchScript)) {
  throw "Arquivo nao encontrado: $watchScript"
}

if ($IntervalMinutes -lt 1) {
  throw 'IntervalMinutes deve ser >= 1.'
}

$taskCmd = "powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$watchScript`" -TunnelName `"$TunnelName`" -Port $Port -FrontendOrigin `"$FrontendOrigin`""

schtasks /Create /F /SC MINUTE /MO $IntervalMinutes /TN $TaskName /TR $taskCmd | Out-Null
Write-Host "Tarefa de monitoramento '$TaskName' instalada (a cada $IntervalMinutes minuto(s))."
