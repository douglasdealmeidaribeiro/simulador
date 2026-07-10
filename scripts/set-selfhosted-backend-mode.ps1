param(
  [Parameter(Mandatory = $true)]
  [string]$PublicApiUrl,

  [switch]$Disable
)

$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
$configPath = Join-Path $root 'assets\api-config.js'

if (-not (Test-Path -LiteralPath $configPath)) {
  throw "Arquivo nao encontrado: $configPath"
}

$content = Get-Content -LiteralPath $configPath -Raw

if ($Disable) {
  $updated = $content -replace "baseUrl:\s*'[^']*'", "baseUrl: ''"
  $updated = $updated -replace 'preferExcelBackend:\s*false', 'preferExcelBackend: true'
  Set-Content -LiteralPath $configPath -Value $updated -Encoding UTF8
  Write-Host 'URL publica removida de assets/api-config.js. O app continua em modo exato e exigira backend acessivel.'
  exit 0
}

$normalized = $PublicApiUrl.Trim().TrimEnd('/')
if (-not ($normalized -match '^https?://')) {
  throw 'Informe uma URL publica valida, por exemplo: https://api.seudominio.com'
}

$escapedUrl = $normalized.Replace('$', '$$')
$updated = $content -replace "baseUrl:\s*'[^']*'", "baseUrl: '$escapedUrl'"
if ($updated -eq $content -and $content -notmatch [regex]::Escape("baseUrl: '$normalized'")) {
  throw 'Nao foi possivel atualizar baseUrl em assets/api-config.js'
}
$updated = $updated -replace 'preferExcelBackend:\s*false', 'preferExcelBackend: true'
Set-Content -LiteralPath $configPath -Value $updated -Encoding UTF8

Write-Host "Modo self-hosted habilitado com API publica: $normalized"
