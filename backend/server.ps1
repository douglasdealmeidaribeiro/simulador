param(
  [int]$Port = 3000,
  [string]$WorkbookPath = '',
  [string]$FrontendOrigin = '*'
)

$ErrorActionPreference = 'Stop'
$RootPath = Split-Path -Parent $PSScriptRoot

if ([string]::IsNullOrWhiteSpace($WorkbookPath)) {
  $workbookFile = Get-ChildItem -LiteralPath $RootPath -Filter '*.xlsm' | Select-Object -First 1
  if ($null -eq $workbookFile) {
    throw 'Nenhum arquivo .xlsm foi encontrado na pasta do projeto.'
  }
  $WorkbookPath = $workbookFile.FullName
}

function Find-Bytes {
  param(
    [byte[]]$Bytes,
    [byte[]]$Pattern
  )

  if ($Bytes.Length -lt $Pattern.Length) {
    return -1
  }

  for ($i = 0; $i -le $Bytes.Length - $Pattern.Length; $i++) {
    $matched = $true
    for ($j = 0; $j -lt $Pattern.Length; $j++) {
      if ($Bytes[$i + $j] -ne $Pattern[$j]) {
        $matched = $false
        break
      }
    }
    if ($matched) {
      return $i
    }
  }

  return -1
}

function Read-ExactBytes {
  param(
    [System.IO.Stream]$Stream,
    [int]$Length
  )

  $buffer = New-Object byte[] $Length
  $offset = 0
  while ($offset -lt $Length) {
    $read = $Stream.Read($buffer, $offset, $Length - $offset)
    if ($read -le 0) {
      break
    }
    $offset += $read
  }

  if ($offset -eq $Length) {
    return $buffer
  }

  $partial = New-Object byte[] $offset
  [Array]::Copy($buffer, $partial, $offset)
  return $partial
}

function Read-HttpRequest {
  param([System.Net.Sockets.TcpClient]$Client)

  $stream = $Client.GetStream()
  $headerBytes = New-Object System.Collections.Generic.List[byte]
  $endMarker = [System.Text.Encoding]::ASCII.GetBytes("`r`n`r`n")
  $readBuffer = New-Object byte[] 1

  while ($true) {
    $read = $stream.Read($readBuffer, 0, 1)
    if ($read -le 0) {
      break
    }
    $headerBytes.Add($readBuffer[0])
    if ($headerBytes.Count -ge 4 -and (Find-Bytes -Bytes $headerBytes.ToArray() -Pattern $endMarker) -ge 0) {
      break
    }
    if ($headerBytes.Count -gt 16384) {
      throw 'Cabeçalho HTTP muito grande.'
    }
  }

  $headerText = [System.Text.Encoding]::ASCII.GetString($headerBytes.ToArray())
  $lines = $headerText -split "`r`n"
  $requestLine = $lines[0] -split ' '
  if ($requestLine.Count -lt 2) {
    throw 'Requisição HTTP inválida.'
  }

  $headers = @{}
  foreach ($line in $lines | Select-Object -Skip 1) {
    if ([string]::IsNullOrWhiteSpace($line)) {
      continue
    }
    $separator = $line.IndexOf(':')
    if ($separator -lt 1) {
      continue
    }
    $name = $line.Substring(0, $separator).Trim().ToLowerInvariant()
    $value = $line.Substring($separator + 1).Trim()
    $headers[$name] = $value
  }

  $contentLength = 0
  if ($headers.ContainsKey('content-length')) {
    $contentLength = [int]$headers['content-length']
  }

  $body = ''
  if ($contentLength -gt 0) {
    $bodyBytes = Read-ExactBytes -Stream $stream -Length $contentLength
    $body = [System.Text.Encoding]::UTF8.GetString($bodyBytes)
  }

  return @{
    Method = $requestLine[0]
    Path = ([Uri]::UnescapeDataString(($requestLine[1] -split '\?')[0]))
    Headers = $headers
    Body = $body
  }
}

function Write-HttpResponse {
  param(
    [System.Net.Sockets.TcpClient]$Client,
    [int]$StatusCode,
    [string]$Reason,
    [hashtable]$Headers,
    [byte[]]$Body
  )

  $stream = $Client.GetStream()
  $allHeaders = [ordered]@{
    'Access-Control-Allow-Origin' = $FrontendOrigin
    'Access-Control-Allow-Methods' = 'GET,POST,OPTIONS'
    'Access-Control-Allow-Headers' = 'Content-Type'
    'Content-Length' = $Body.Length
    'Connection' = 'close'
  }

  foreach ($key in $Headers.Keys) {
    $allHeaders[$key] = $Headers[$key]
  }

  $headerText = "HTTP/1.1 $StatusCode $Reason`r`n"
  foreach ($key in $allHeaders.Keys) {
    $headerText += "${key}: $($allHeaders[$key])`r`n"
  }
  $headerText += "`r`n"

  $headerBytes = [System.Text.Encoding]::ASCII.GetBytes($headerText)
  $stream.Write($headerBytes, 0, $headerBytes.Length)
  if ($Body.Length -gt 0) {
    $stream.Write($Body, 0, $Body.Length)
  }
  $stream.Flush()
  $Client.Close()
}

function Send-Json {
  param(
    [System.Net.Sockets.TcpClient]$Client,
    [int]$StatusCode,
    [string]$Reason,
    [object]$Payload
  )

  $json = $Payload | ConvertTo-Json -Compress -Depth 10
  $body = [System.Text.Encoding]::UTF8.GetBytes($json)
  Write-HttpResponse -Client $Client -StatusCode $StatusCode -Reason $Reason -Headers @{
    'Content-Type' = 'application/json; charset=utf-8'
  } -Body $body
}

function Invoke-Simulation {
  param([string]$Body)

  $scriptPath = Join-Path $PSScriptRoot 'simulate-excel.ps1'
  $resolvedWorkbook = (Resolve-Path -LiteralPath $WorkbookPath).Path
  $workDir = Join-Path ([System.IO.Path]::GetTempPath()) 'simulador-excel'
  [System.IO.Directory]::CreateDirectory($workDir) | Out-Null

  $id = '{0}-{1}' -f ([DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds()), ([Guid]::NewGuid().ToString('N'))
  $inputPath = Join-Path $workDir "$id.json"
  $outputPath = Join-Path $workDir "$id.xlsm"
  $stdoutPath = Join-Path $workDir "$id.out.txt"
  $stderrPath = Join-Path $workDir "$id.err.txt"

  try {
    $payload = $Body | ConvertFrom-Json
    if ($null -eq $payload -or $null -eq $payload.updates -or $payload.updates -is [array]) {
      throw 'Envie um JSON no formato { "updates": { "C3": 0.0197 } }.'
    }

    [System.IO.File]::WriteAllText($inputPath, $Body, [System.Text.UTF8Encoding]::new($false))
    $arguments = @(
      '-STA'
      '-NoProfile'
      '-ExecutionPolicy'
      'Bypass'
      '-File'
      "`"$scriptPath`""
      '-WorkbookPath'
      "`"$resolvedWorkbook`""
      '-InputJsonPath'
      "`"$inputPath`""
      '-OutputWorkbookPath'
      "`"$outputPath`""
    ) -join ' '

    $process = Start-Process -FilePath 'powershell.exe' -ArgumentList $arguments -Wait -PassThru -RedirectStandardOutput $stdoutPath -RedirectStandardError $stderrPath
    if ($process.ExitCode -ne 0) {
      $details = @()
      if (Test-Path -LiteralPath $stderrPath) {
        $details += Get-Content -LiteralPath $stderrPath -Raw
      }
      if (Test-Path -LiteralPath $stdoutPath) {
        $details += Get-Content -LiteralPath $stdoutPath -Raw
      }
      throw "Excel finalizou com código $($process.ExitCode). $($details -join ' ')"
    }

    if (-not (Test-Path -LiteralPath $outputPath)) {
      throw 'A planilha calculada não foi gerada.'
    }

    return [System.IO.File]::ReadAllBytes($outputPath)
  } finally {
    Remove-Item -LiteralPath $inputPath -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $outputPath -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $stdoutPath -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $stderrPath -ErrorAction SilentlyContinue
  }
}

$listener = [System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Parse('127.0.0.1'), $Port)
$listener.Start()

Write-Host "Simulador Excel API em http://127.0.0.1:$Port"
Write-Host 'Pressione Ctrl+C para encerrar.'

try {
  while ($true) {
    $client = $listener.AcceptTcpClient()
    try {
      $request = Read-HttpRequest -Client $client

      if ($request.Method -eq 'OPTIONS') {
        Write-HttpResponse -Client $client -StatusCode 204 -Reason 'No Content' -Headers @{} -Body ([byte[]]@())
        continue
      }

      if ($request.Method -eq 'GET' -and $request.Path -eq '/health') {
        Send-Json -Client $client -StatusCode 200 -Reason 'OK' -Payload @{ ok = $true }
        continue
      }

      if ($request.Method -eq 'POST' -and $request.Path -eq '/api/simular') {
        try {
          $file = Invoke-Simulation -Body $request.Body
          Write-HttpResponse -Client $client -StatusCode 200 -Reason 'OK' -Headers @{
            'Content-Type' = 'application/vnd.ms-excel.sheet.macroEnabled.12'
            'Content-Disposition' = 'attachment; filename="simulador-modelagem-calculado.xlsm"'
          } -Body $file
        } catch {
          Send-Json -Client $client -StatusCode 500 -Reason 'Internal Server Error' -Payload @{ error = $_.Exception.Message }
        }
        continue
      }

      Send-Json -Client $client -StatusCode 404 -Reason 'Not Found' -Payload @{ error = 'Rota não encontrada.' }
    } catch {
      try {
        Send-Json -Client $client -StatusCode 500 -Reason 'Internal Server Error' -Payload @{ error = $_.Exception.Message }
      } catch {
        $client.Close()
      }
    }
  }
} finally {
  $listener.Stop()
}
