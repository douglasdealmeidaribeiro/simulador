param(
  [string]$TaskName = 'SimuladorExactSelfHostedMonitor'
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$launcherVbs = Join-Path $root 'scripts\launch-selfhosted-monitor-hidden.vbs'

schtasks /Delete /F /TN $TaskName | Out-Null
Write-Host "Tarefa de monitoramento '$TaskName' removida."

if (Test-Path -LiteralPath $launcherVbs) {
  Remove-Item -LiteralPath $launcherVbs -Force -ErrorAction SilentlyContinue
}
