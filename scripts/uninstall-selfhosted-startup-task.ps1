param(
  [string]$TaskName = 'SimuladorExactSelfHosted'
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$launcherVbs = Join-Path $root 'scripts\launch-selfhosted-startup-hidden.vbs'

if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
  Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
  Write-Host "Tarefa '$TaskName' removida."
} else {
  Write-Host "Tarefa '$TaskName' nao encontrada."
}

$runKey = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run'
if ((Get-ItemProperty -Path $runKey -ErrorAction SilentlyContinue).PSObject.Properties.Name -contains $TaskName) {
  Remove-ItemProperty -Path $runKey -Name $TaskName -ErrorAction SilentlyContinue
  Write-Host "Fallback HKCU\\Run '$TaskName' removido."
}

if (Test-Path -LiteralPath $launcherVbs) {
  Remove-Item -LiteralPath $launcherVbs -Force -ErrorAction SilentlyContinue
}
