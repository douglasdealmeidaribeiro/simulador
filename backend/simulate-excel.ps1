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

$resolvedWorkbook = (Resolve-Path -LiteralPath $WorkbookPath).Path
$payload = Get-Content -LiteralPath $InputJsonPath -Raw | ConvertFrom-Json

Copy-Item -LiteralPath $resolvedWorkbook -Destination $OutputWorkbookPath -Force

$excel = $null
$workbook = $null
$simulador = $null
$controle = $null

try {
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $excel.AskToUpdateLinks = $false
  $excel.AutomationSecurity = 1
  $excel.EnableEvents = $false

  $workbook = Invoke-WithRetry { $excel.Workbooks.Open($OutputWorkbookPath) }
  $simulador = Invoke-WithRetry { $workbook.Worksheets.Item('SIMULADOR') }
  $controle = Invoke-WithRetry { $workbook.Worksheets.Item('Controle') }

  foreach ($property in $payload.updates.PSObject.Properties) {
    $cell = [string]$property.Name
    Assert-CellReference -Reference $cell
    $value = $property.Value
    if ($value -is [double] -or $value -is [int] -or $value -is [decimal]) {
      $simulador.Range($cell).Value2 = [double]$value
    } else {
      $simulador.Range($cell).Value2 = [string]$value
    }
  }

  $excel.CalculateFullRebuild()

  try {
    $excel.Run("'" + $workbook.Name + "'!SimularESDigital")
  } catch {
    $controle.Range('M3').Value2 = 15000
    $goalCell = $controle.Range('M13')
    $changingCell = $controle.Range('M3')
    [void]$goalCell.GoalSeek(0.1, $changingCell)
  }

  $excel.CalculateFullRebuild()
  $workbook.Save()

  $summary = [ordered]@{
    annual = $simulador.Range('G3').Value2
    monthly = $simulador.Range('G4').Value2
    target = $controle.Range('M13').Value2
  }
  $summary | ConvertTo-Json -Compress
} finally {
  if ($null -ne $workbook) {
    try {
      $workbook.Close($true)
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
  Release-ComObject $controle
  Release-ComObject $simulador
  Release-ComObject $workbook
  Release-ComObject $excel
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
