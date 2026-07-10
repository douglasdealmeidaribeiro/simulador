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
$launcherVbs = Join-Path $root 'scripts\launch-selfhosted-monitor-hidden.vbs'

if (-not (Test-Path -LiteralPath $watchScript)) {
  throw "Arquivo nao encontrado: $watchScript"
}

if ($IntervalMinutes -lt 1) {
  throw 'IntervalMinutes deve ser >= 1.'
}

$safeWatchScript = $watchScript.Replace('"', '""')
$safeTunnelName = $TunnelName.Replace('"', '""')
$safeFrontendOrigin = $FrontendOrigin.Replace('"', '""')
$psCommand = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"`"$safeWatchScript`"`" -TunnelName `"`"$safeTunnelName`"`" -Port $Port -FrontendOrigin `"`"$safeFrontendOrigin`"`""
$vbsContent = @"
Set shell = CreateObject("WScript.Shell")
shell.Run "$psCommand", 0, False
"@
Set-Content -LiteralPath $launcherVbs -Value $vbsContent -Encoding ASCII

$taskCmd = "wscript.exe `"$launcherVbs`""

schtasks /Create /F /SC MINUTE /MO $IntervalMinutes /TN $TaskName /TR $taskCmd | Out-Null
Write-Host "Tarefa de monitoramento '$TaskName' instalada (a cada $IntervalMinutes minuto(s))."
