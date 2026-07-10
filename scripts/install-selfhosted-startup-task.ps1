param(
  [string]$TaskName = 'SimuladorExactSelfHosted',
  [string]$TunnelName = 'simulador-api',
  [int]$Port = 3000,
  [string]$FrontendOrigin = '*'
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$startScript = Join-Path $root 'scripts\start-selfhosted-exact-background.ps1'
$launcherVbs = Join-Path $root 'scripts\launch-selfhosted-startup-hidden.vbs'

if (-not (Test-Path -LiteralPath $startScript)) {
  throw "Arquivo nao encontrado: $startScript"
}

$safeStartScript = $startScript.Replace('"', '""')
$safeTunnelName = $TunnelName.Replace('"', '""')
$safeFrontendOrigin = $FrontendOrigin.Replace('"', '""')
$psCommand = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"`"$safeStartScript`"`" -TunnelName `"`"$safeTunnelName`"`" -Port $Port -FrontendOrigin `"`"$safeFrontendOrigin`"`""
$vbsContent = @"
Set shell = CreateObject("WScript.Shell")
shell.Run "$psCommand", 0, False
"@
Set-Content -LiteralPath $launcherVbs -Value $vbsContent -Encoding ASCII

$action = New-ScheduledTaskAction -Execute 'wscript.exe' -Argument "`"$launcherVbs`""
$trigger = New-ScheduledTaskTrigger -AtLogOn -User "$env:USERDOMAIN\$env:USERNAME"
$principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType Interactive -RunLevel Highest
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

try {
  Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Force | Out-Null
  Write-Host "Tarefa '$TaskName' instalada para iniciar no logon do usuario $env:USERNAME."
  exit 0
} catch {
  $accessDenied = $_.Exception.Message -like '*0x80070005*' -or $_.Exception.Message -like '*Acesso negado*'
  if (-not $accessDenied) {
    throw
  }
}

# Fallback sem privilegio de administrador: inicializacao do usuario atual via HKCU\Run.
$runKey = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run'
$runName = $TaskName
$runValue = "wscript.exe `"$launcherVbs`""
Set-ItemProperty -Path $runKey -Name $runName -Value $runValue -Force
Write-Host "Sem permissao para Scheduled Task. Criado fallback de inicializacao em HKCU\\Run com nome '$runName'."
