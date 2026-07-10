param(
  [Parameter(Mandatory = $true)]
  [string]$WorkbookPath,

  [Parameter(Mandatory = $true)]
  [string]$InputJsonPath,

  [Parameter(Mandatory = $true)]
  [string]$OutputWorkbookPath
)

$ErrorActionPreference = 'Stop'

function Release-ComObject {
  param([object]$Object)
  try {
    if ($null -ne $Object -and [System.Runtime.InteropServices.Marshal]::IsComObject($Object)) {
      [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($Object)
    }
  } catch {
    # Best effort cleanup.
  }
}

function Assert-CellReference {
  param([string]$Reference)
  if ($Reference -notmatch '^[A-Z]{1,3}[1-9][0-9]{0,6}$') {
    throw "Referência de célula inválida: $Reference"
  }
}

function Invoke-WithRetry {
  param(
    [scriptblock]$Action,
    [int]$Attempts = 20,
    [int]$DelayMilliseconds = 500
  )

  $lastError = $null
  for ($attempt = 1; $attempt -le $Attempts; $attempt++) {
    try {
      $result = & $Action
      if ($null -ne $result) {
        return $result
      }
    } catch {
      $lastError = $_
    }
    Start-Sleep -Milliseconds $DelayMilliseconds
  }

  if ($null -ne $lastError) {
    throw $lastError
  }
  throw 'O Excel não retornou o objeto esperado.'
}

function Invoke-ComActionWithRetry {
  param(
    [scriptblock]$Action,
    [int]$Attempts = 20,
    [int]$DelayMilliseconds = 500
  )

  $lastError = $null
  for ($attempt = 1; $attempt -le $Attempts; $attempt++) {
    try {
      & $Action
      return
    } catch {
      $lastError = $_
    }
    Start-Sleep -Milliseconds $DelayMilliseconds
  }

  if ($null -ne $lastError) {
    throw $lastError
  }
  throw 'A automacao do Excel nao concluiu a operacao.'
}

function Normalize-Ascii {
  param([string]$Text)
  if ($null -eq $Text) {
    return ''
  }
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

$resolvedWorkbook = (Resolve-Path -LiteralPath $WorkbookPath).Path
$payload = Get-Content -LiteralPath $InputJsonPath -Raw | ConvertFrom-Json

$calcWorkbookPath = [System.IO.Path]::Combine(
  [System.IO.Path]::GetDirectoryName($OutputWorkbookPath),
  [System.IO.Path]::GetFileNameWithoutExtension($OutputWorkbookPath) + '-calc.xlsm'
)

Copy-Item -LiteralPath $resolvedWorkbook -Destination $calcWorkbookPath -Force

$excel = $null
$workbook = $null
$simulador = $null
$controle = $null
$orcamentoMensal = $null
$mensalWorkbook = $null
$mensalSheet = $null
$usedRange = $null

try {
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $excel.AskToUpdateLinks = $false
  $excel.AutomationSecurity = 1
  $excel.EnableEvents = $false

  $workbook = Invoke-WithRetry { $excel.Workbooks.Open($calcWorkbookPath) }
  $simulador = Invoke-WithRetry { $workbook.Worksheets.Item('SIMULADOR') }
  $controle = Invoke-WithRetry { $workbook.Worksheets.Item('Controle') }

  foreach ($property in $payload.updates.PSObject.Properties) {
    $cell = [string]$property.Name
    Assert-CellReference -Reference $cell
    $value = $property.Value
    if ($value -is [double] -or $value -is [int] -or $value -is [decimal]) {
      Invoke-ComActionWithRetry { $simulador.Range($cell).Value2 = [double]$value }
    } else {
      Invoke-ComActionWithRetry { $simulador.Range($cell).Value2 = [string]$value }
    }
  }

  Invoke-ComActionWithRetry { $excel.CalculateFullRebuild() }

  try {
    Invoke-ComActionWithRetry { $excel.Run("'" + $workbook.Name + "'!SimularESDigital") }
  } catch {
    Invoke-ComActionWithRetry { $controle.Range('M3').Value2 = 15000 }
    $goalCell = $controle.Range('M13')
    $changingCell = $controle.Range('M3')
    Invoke-ComActionWithRetry { [void]$goalCell.GoalSeek(0.1, $changingCell) }
  }

  Invoke-ComActionWithRetry { $excel.CalculateFullRebuild() }

  $orcamentoMensal = Get-WorksheetLike -Workbook $workbook -Pattern '(mensal)'
  Invoke-ComActionWithRetry { $orcamentoMensal.Copy() }
  $mensalWorkbook = Invoke-WithRetry { $excel.ActiveWorkbook }
  $mensalSheet = Invoke-WithRetry { $mensalWorkbook.Worksheets.Item(1) }
  $usedRange = $mensalSheet.UsedRange
  # Converte formulas em valores para entregar uma planilha final fechada e fiel ao resultado pos-simulacao.
  Invoke-ComActionWithRetry { $usedRange.Value2 = $usedRange.Value2 }
  Invoke-ComActionWithRetry { [void]$mensalWorkbook.SaveAs($OutputWorkbookPath, 51) }
  Invoke-ComActionWithRetry { $mensalWorkbook.Close($true) }
  $mensalWorkbook = $null

  $summary = [ordered]@{
    annual = $simulador.Range('G3').Value2
    monthly = $simulador.Range('G4').Value2
    price = $controle.Range('M3').Value2
    target = $controle.Range('M13').Value2
  }
  $summary | ConvertTo-Json -Compress
} finally {
  if ($null -ne $mensalWorkbook) {
    try {
      $mensalWorkbook.Close($true)
    } catch {
      # Best effort cleanup.
    }
  }
  if ($null -ne $workbook) {
    try {
      $workbook.Close($false)
    } catch {
      # Best effort cleanup.
    }
  }
  if ($null -ne $excel) {
    try {
      $excel.Quit()
    } catch {
      # Best effort cleanup.
    }
  }
  Release-ComObject $usedRange
  Release-ComObject $mensalSheet
  Release-ComObject $mensalWorkbook
  Release-ComObject $orcamentoMensal
  Release-ComObject $controle
  Release-ComObject $simulador
  Release-ComObject $workbook
  Release-ComObject $excel
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
  if (Test-Path -LiteralPath $calcWorkbookPath) {
    try {
      Remove-Item -LiteralPath $calcWorkbookPath -Force
    } catch {
      # Best effort cleanup.
    }
  }
}
