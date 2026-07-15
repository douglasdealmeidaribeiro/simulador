param(
  [string]$ProjectPath = 'C:\simulador',
  [string]$RepoUrl = 'https://github.com/douglasdealmeidaribeiro/simulador.git',
  [int]$Port = 80,
  [string]$BindAddress = '0.0.0.0',
  [string]$PublicHost = ''
)

$ErrorActionPreference = 'Stop'

function Assert-Command {
  param([string]$Name)
  if (-not (Get-Command $Name -ErrorAction SilentlyContinue)) {
    throw "Comando nao encontrado: $Name. Instale antes de continuar."
  }
}

function Test-ExcelCom {
  $excel = $null
  try {
    $excel = New-Object -ComObject Excel.Application
    $version = $excel.Version
    Write-Host "Excel COM OK. Versao: $version"
  } finally {
    if ($null -ne $excel) {
      $excel.Quit()
      [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
  }
}

Assert-Command git

if (Test-Path -LiteralPath (Join-Path $ProjectPath '.git')) {
  Write-Host "Atualizando repositorio em $ProjectPath..."
  git -C $ProjectPath pull --ff-only
} else {
  if (-not (Test-Path -LiteralPath $ProjectPath)) {
    New-Item -ItemType Directory -Path $ProjectPath -Force | Out-Null
  }
  Write-Host "Clonando repositorio em $ProjectPath..."
  git clone $RepoUrl $ProjectPath
}

Test-ExcelCom

$ruleName = "Simulador Excel HTTP $Port"
if (-not (Get-NetFirewallRule -DisplayName $ruleName -ErrorAction SilentlyContinue)) {
  New-NetFirewallRule -DisplayName $ruleName -Direction Inbound -Action Allow -Protocol TCP -LocalPort $Port | Out-Null
  Write-Host "Firewall liberado na porta $Port."
} else {
  Write-Host "Regra de firewall ja existe: $ruleName"
}

$serverScript = Join-Path $ProjectPath 'backend\server.ps1'
$taskName = 'Simulador Excel Backend'
$arguments = "-STA -NoProfile -ExecutionPolicy Bypass -File `"$serverScript`" -Port $Port -BindAddress $BindAddress -FrontendOrigin `"*`""
$action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument $arguments
$trigger = New-ScheduledTaskTrigger -AtLogOn
$principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType Interactive -RunLevel Highest

Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Principal $principal -Force | Out-Null
Write-Host "Tarefa agendada criada/atualizada: $taskName"

Start-Process -FilePath 'powershell.exe' -ArgumentList $arguments
Write-Host "Backend iniciado."

if ([string]::IsNullOrWhiteSpace($PublicHost)) {
  $PublicHost = $env:COMPUTERNAME
}

$portSuffix = if ($Port -eq 80) { '' } else { ":$Port" }
Write-Host "Acesse: http://$PublicHost$portSuffix/"
