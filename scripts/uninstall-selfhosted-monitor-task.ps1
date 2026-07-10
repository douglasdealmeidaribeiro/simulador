param(
  [string]$TaskName = 'SimuladorExactSelfHostedMonitor'
)

$ErrorActionPreference = 'Stop'

schtasks /Delete /F /TN $TaskName | Out-Null
Write-Host "Tarefa de monitoramento '$TaskName' removida."
