param(
  [string]$ScenarioName = 'shelters_desligado'
)

$ErrorActionPreference = 'Stop'

function Release-ComObject {
  param([object]$Object)
  try {
    if ($null -ne $Object -and [System.Runtime.InteropServices.Marshal]::IsComObject($Object)) {
      [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($Object)
    }
  } catch {
    # best effort
  }
}

function Open-WorkbookRetry {
  param(
    [Parameter(Mandatory = $true)]$Excel,
    [Parameter(Mandatory = $true)][string]$Path
  )

  for ($attempt = 0; $attempt -lt 16; $attempt++) {
    try {
      $wb = $Excel.Workbooks.Open($Path)
      if ($null -ne $wb) {
        return $wb
      }
    } catch {
      Start-Sleep -Milliseconds 250
    }
  }

  throw "Nao foi possivel abrir workbook temporario: $Path"
}

function To-DoubleOrNull {
  param([object]$Value)

  if ($Value -is [double] -or $Value -is [float] -or $Value -is [int] -or $Value -is [long] -or $Value -is [decimal]) {
    return [double]$Value
  }
  if ($null -eq $Value) {
    return $null
  }

  $asString = [string]$Value
  try {
    return [double]::Parse($asString, [System.Globalization.CultureInfo]::InvariantCulture)
  } catch {
    try {
      return [double]::Parse($asString, [System.Globalization.CultureInfo]::GetCultureInfo('pt-BR'))
    } catch {
      return $null
    }
  }
}

function Normalize-Ascii {
  param([string]$Text)
  if ($null -eq $Text) { return '' }
  $normalized = $Text.Normalize([Text.NormalizationForm]::FormD)
  $ascii = -join ($normalized.ToCharArray() | Where-Object { [Globalization.CharUnicodeInfo]::GetUnicodeCategory($_) -ne [Globalization.UnicodeCategory]::NonSpacingMark })
  return $ascii.ToLowerInvariant()
}

function Get-WorksheetLike {
  param(
    [Parameter(Mandatory = $true)]$Workbook,
    [Parameter(Mandatory = $true)][string]$Pattern
  )

  $target = Normalize-Ascii $Pattern
  foreach ($ws in $Workbook.Worksheets) {
    $name = [string]$ws.Name
    if ((Normalize-Ascii $name) -like "*$target*") {
      return $ws
    }
  }

  throw "Aba nao encontrada com padrao: $Pattern"
}

function Scenario-Updates {
  param([string]$Name)

  switch ($Name) {
    'base' { return @{} }
    'shelters_desligado' { return @{ C15 = 'NÃO' } }
    'variavel_desligado' { return @{ C25 = 'NÃO' } }
    'riscos_altos' { return @{ C3 = 0.05; C4 = 0.04 } }
    'ajuste_orcamento_misto' { return @{ G14 = 0.15; G32 = -0.08; G66 = 0.12 } }
    default { throw "Cenario desconhecido: $Name" }
  }
}

$root = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$webScript = Join-Path $root 'scripts\simulate_web.js'
$excelScript = Join-Path $root 'backend\simulate-excel.ps1'
$workbookPath = Get-ChildItem -LiteralPath $root -Filter *.xlsm | Select-Object -First 1 -ExpandProperty FullName

if (-not (Test-Path -LiteralPath $webScript)) { throw "Nao encontrado: $webScript" }
if (-not (Test-Path -LiteralPath $excelScript)) { throw "Nao encontrado: $excelScript" }
if (-not $workbookPath) { throw 'Nenhum .xlsm encontrado na raiz do projeto.' }

$updates = Scenario-Updates -Name $ScenarioName
$refs = 1..149 | ForEach-Object { "D$_" }
$debugArg = $refs -join ','

$tmpIn = Join-Path $env:TEMP ("sim-diff-in-" + [Guid]::NewGuid().ToString() + '.json')
$tmpOut = Join-Path $env:TEMP ("sim-diff-out-" + [Guid]::NewGuid().ToString() + '.xlsm')

$excel = $null
$wb = $null
$sim = $null
$ctrl = $null
$orig = $null

try {
  $updates | ConvertTo-Json -Compress | Set-Content -LiteralPath $tmpIn -Encoding UTF8
  $webJson = & node $webScript $tmpIn $debugArg
  if ($LASTEXITCODE -ne 0) {
    throw 'Falha na simulacao web.'
  }
  $web = $webJson | ConvertFrom-Json

  @{ updates = $updates } | ConvertTo-Json -Compress | Set-Content -LiteralPath $tmpIn -Encoding UTF8
  $excelJson = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $excelScript -WorkbookPath $workbookPath -InputJsonPath $tmpIn -OutputWorkbookPath $tmpOut
  if ($LASTEXITCODE -ne 0) {
    throw 'Falha na simulacao Excel.'
  }
  $excelSummary = $excelJson | ConvertFrom-Json

  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $wb = Open-WorkbookRetry -Excel $excel -Path $tmpOut
  $sim = $wb.Worksheets.Item('SIMULADOR')
  $ctrl = $wb.Worksheets.Item('Controle')
  $orig = Get-WorksheetLike -Workbook $wb -Pattern 'orcamento (origem)'

  $diffs = New-Object System.Collections.Generic.List[object]
  foreach ($ref in $refs) {
    $webVal = To-DoubleOrNull $web.debugOrigem.$ref
    $excelVal = To-DoubleOrNull $orig.Range($ref).Value2
    if ($null -eq $webVal -or $null -eq $excelVal) {
      continue
    }
    $delta = $webVal - $excelVal
    if ([math]::Abs($delta) -gt 0.0001) {
      $diffs.Add([pscustomobject]@{
        ref = $ref
        web = $webVal
        excel = $excelVal
        delta = $delta
      }) | Out-Null
    }
  }

  Write-Host "Cenario: $ScenarioName"
  Write-Host "Resultado anual: web=$($web.annual) excelResumo=$($excelSummary.annual) excelCelula=$([double]$sim.Range('G3').Value2)"
  Write-Host "Resultado mensal: web=$($web.monthly) excelCelula=$([double]$sim.Range('G4').Value2)"
  Write-Host "Meta: web=$($web.target) excelResumo=$($excelSummary.target) excelCelula=$([double]$ctrl.Range('M13').Value2)"
  Write-Host ''

  if ($diffs.Count -eq 0) {
    Write-Host 'Nenhuma diferenca relevante encontrada em Orçamento (Origem)!D1:D149.'
  } else {
    $diffs |
      Sort-Object { [math]::Abs($_.delta) } -Descending |
      Select-Object -First 60 |
      Format-Table -AutoSize
  }
} finally {
  if ($null -ne $wb) {
    try { $wb.Close($false) } catch { }
  }
  if ($null -ne $excel) {
    try { $excel.Quit() } catch { }
  }

  Release-ComObject $orig
  Release-ComObject $ctrl
  Release-ComObject $sim
  Release-ComObject $wb
  Release-ComObject $excel

  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()

  if (Test-Path -LiteralPath $tmpIn) { Remove-Item -LiteralPath $tmpIn -Force }
  if (Test-Path -LiteralPath $tmpOut) { Remove-Item -LiteralPath $tmpOut -Force }
}
