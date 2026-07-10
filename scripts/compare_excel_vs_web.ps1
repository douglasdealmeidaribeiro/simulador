param(
  [switch]$StopOnFirstError
)

$ErrorActionPreference = 'Stop'

$root = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$webScript = Join-Path $root 'scripts\simulate_web.js'
$excelScript = Join-Path $root 'backend\simulate-excel.ps1'
$workbookPath = Get-ChildItem -LiteralPath $root -Filter *.xlsm | Select-Object -First 1 -ExpandProperty FullName

if (-not (Test-Path -LiteralPath $webScript)) {
  throw "Script web nao encontrado: $webScript"
}
if (-not (Test-Path -LiteralPath $excelScript)) {
  throw "Script Excel nao encontrado: $excelScript"
}
if (-not $workbookPath) {
  throw 'Nenhum arquivo .xlsm encontrado na raiz do projeto.'
}

function Invoke-WebSimulation {
  param([hashtable]$Updates)

  $tmpIn = Join-Path $env:TEMP ("sim-web-" + [Guid]::NewGuid().ToString() + '.json')
  try {
    $Updates | ConvertTo-Json -Compress | Set-Content -LiteralPath $tmpIn -Encoding UTF8
    $json = & node $webScript $tmpIn
    if ($LASTEXITCODE -ne 0) {
      throw "Falha ao executar simulacao web."
    }
    return ($json | ConvertFrom-Json)
  } finally {
    if (Test-Path -LiteralPath $tmpIn) {
      Remove-Item -LiteralPath $tmpIn -Force
    }
  }
}

function Invoke-ExcelSimulation {
  param([hashtable]$Updates)

  $tmpIn = Join-Path $env:TEMP ("sim-excel-in-" + [Guid]::NewGuid().ToString() + '.json')
  $tmpOut = Join-Path $env:TEMP ("sim-excel-out-" + [Guid]::NewGuid().ToString() + '.xlsm')
  try {
    @{ updates = $Updates } | ConvertTo-Json -Compress | Set-Content -LiteralPath $tmpIn -Encoding UTF8
    $json = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $excelScript -WorkbookPath $workbookPath -InputJsonPath $tmpIn -OutputWorkbookPath $tmpOut
    if ($LASTEXITCODE -ne 0) {
      throw 'Falha ao executar simulacao Excel.'
    }
    return ($json | ConvertFrom-Json)
  } finally {
    if (Test-Path -LiteralPath $tmpIn) {
      Remove-Item -LiteralPath $tmpIn -Force
    }
    if (Test-Path -LiteralPath $tmpOut) {
      Remove-Item -LiteralPath $tmpOut -Force
    }
  }
}

function New-Scenario {
  param(
    [string]$Name,
    [hashtable]$Updates
  )
  return [pscustomobject]@{ Name = $Name; Updates = $Updates }
}

$scenarios = @(
  (New-Scenario -Name 'base' -Updates @{}),
  (New-Scenario -Name 'shelters_desligado' -Updates @{ C15 = 'NÃO' }),
  (New-Scenario -Name 'variavel_desligado' -Updates @{ C25 = 'NÃO' }),
  (New-Scenario -Name 'riscos_altos' -Updates @{ C3 = 0.05; C4 = 0.04 }),
  (New-Scenario -Name 'ajuste_orcamento_misto' -Updates @{ G14 = 0.15; G32 = -0.08; G66 = 0.12 })
)

$rows = New-Object System.Collections.Generic.List[object]

foreach ($scenario in $scenarios) {
  try {
    Write-Host "Executando cenario: $($scenario.Name)"
    $web = Invoke-WebSimulation -Updates $scenario.Updates
    $excel = Invoke-ExcelSimulation -Updates $scenario.Updates

    $deltaAnnual = [double]$web.annual - [double]$excel.annual
    $deltaMonthly = [double]$web.monthly - [double]$excel.monthly
    $deltaTarget = [double]$web.target - [double]$excel.target

    $rows.Add([pscustomobject]@{
      cenario = $scenario.Name
      web_anual = [double]$web.annual
      excel_anual = [double]$excel.annual
      delta_anual = $deltaAnnual
      web_mensal = [double]$web.monthly
      excel_mensal = [double]$excel.monthly
      delta_mensal = $deltaMonthly
      web_meta = [double]$web.target
      excel_meta = [double]$excel.target
      delta_meta = $deltaTarget
    }) | Out-Null
  } catch {
    $rows.Add([pscustomobject]@{
      cenario = $scenario.Name
      erro = $_.Exception.Message
    }) | Out-Null
    if ($StopOnFirstError) {
      throw
    }
  }
}

$rows | Format-Table -AutoSize

Write-Host ''
Write-Host 'JSON:'
$rows | ConvertTo-Json -Depth 4
