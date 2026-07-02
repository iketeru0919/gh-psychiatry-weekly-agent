param(
  [int]$Year = 0,
  [int]$Month = 0,
  [string]$ConfigPath = "",
  [string]$BillingWorkbook = "",
  [string]$OutputPath = "",
  [switch]$Interactive
)

$ErrorActionPreference = "Stop"
$toolDir = $PSScriptRoot
$portableToolDir = Join-Path $toolDir "tool_runtime"
$workspaceRoot = Split-Path (Split-Path $toolDir -Parent) -Parent
$workspaceToolDir = Join-Path $workspaceRoot "work\billing_provider_check"

if (Test-Path (Join-Path $portableToolDir "analyze_check.py")) {
  $workToolDir = $portableToolDir
} else {
  $workToolDir = $workspaceToolDir
}

if ([string]::IsNullOrWhiteSpace($ConfigPath)) {
  $ConfigPath = Join-Path $toolDir "config_auto.json"
}

$python = "C:\Users\k-takayama\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
if (-not (Test-Path $python)) {
  throw "Python runtime not found: $python"
}
if (-not (Test-Path (Join-Path $workToolDir "analyze_check.py"))) {
  throw "analyze_check.py not found. Keep the tool_runtime folder together with run_check.ps1."
}
if (-not (Test-Path (Join-Path $workToolDir "build_report.py"))) {
  throw "build_report.py not found. Keep the tool_runtime folder together with run_check.ps1."
}

$jsonPath = Join-Path $workToolDir "latest_result.json"
$effectiveConfigPath = Join-Path $workToolDir "effective_config.json"

$config = Get-Content -LiteralPath $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json

if ($Interactive) {
  Write-Host ""
  Write-Host "Billing / provider check"
  Write-Host "------------------------"
  Write-Host "Press Enter to use the default value shown in brackets."
  Write-Host ""

  $defaultYear = [int]$config.target_year
  $yearInput = Read-Host ("Target year [{0}]" -f $defaultYear)
  if (-not [string]::IsNullOrWhiteSpace($yearInput)) {
    $Year = [int]$yearInput
  }

  $defaultMonth = [int]$config.target_month
  $monthInput = Read-Host ("Target month 1-12 [{0}]" -f $defaultMonth)
  if (-not [string]::IsNullOrWhiteSpace($monthInput)) {
    $Month = [int]$monthInput
  }

  $billingCandidates = Get-ChildItem -LiteralPath $toolDir -File |
    Where-Object {
      ($_.Extension -in @(".xlsx", ".xlsm")) -and
      $_.Name -notlike "~$*" -and
      $_.Name -notlike "billing_provider_check_*" -and
      $_.Name -notlike "*.inspect.*" -and
      $_.Length -gt 100KB
    } |
    Sort-Object LastWriteTime -Descending

  if ($billingCandidates.Count -gt 0) {
    Write-Host ""
    Write-Host "Billing workbook candidates in this folder:"
    for ($i = 0; $i -lt $billingCandidates.Count; $i++) {
      Write-Host ("  {0}: {1}" -f ($i + 1), $billingCandidates[$i].Name)
    }
    Write-Host "  0: Use billing_workbook in config_auto.json"
    Write-Host ""
    $choice = Read-Host "Select number, or paste full path, or press Enter for config"
    if (-not [string]::IsNullOrWhiteSpace($choice)) {
      $parsed = 0
      if ([int]::TryParse($choice, [ref]$parsed)) {
        if ($parsed -gt 0 -and $parsed -le $billingCandidates.Count) {
          $BillingWorkbook = $billingCandidates[$parsed - 1].FullName
        }
      } else {
        $BillingWorkbook = $choice.Trim('"')
      }
    }
  } else {
    Write-Host ""
    Write-Host "No billing workbook found in this folder."
    $pathInput = Read-Host "Paste billing workbook path, or press Enter for config"
    if (-not [string]::IsNullOrWhiteSpace($pathInput)) {
      $BillingWorkbook = $pathInput.Trim('"')
    }
  }
}

if ($Year -gt 0) {
  $config.target_year = $Year
}
if ($Month -gt 0) {
  $config.target_month = $Month
}
if (-not [string]::IsNullOrWhiteSpace($BillingWorkbook)) {
  $config.billing_workbook = $BillingWorkbook
}
if (-not $config.target_year -or -not $config.target_month) {
  throw "target_year and target_month are required. Specify them in config_auto.json or pass -Year and -Month."
}
if ([string]::IsNullOrWhiteSpace($OutputPath)) {
  $OutputPath = Join-Path $toolDir ("billing_provider_check_{0}_{1:00}.xlsx" -f [int]$config.target_year, [int]$config.target_month)
}

$monthLabel = ("{0}{1}" -f [int]$config.target_month, [char]0x6708)
$config | Add-Member -NotePropertyName "target_month_label" -NotePropertyValue $monthLabel -Force
$config | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $effectiveConfigPath -Encoding UTF8

& $python (Join-Path $workToolDir "analyze_check.py") $effectiveConfigPath $jsonPath
if ($LASTEXITCODE -ne 0) {
  throw "Python analyze failed. ExitCode=$LASTEXITCODE"
}
& $python (Join-Path $workToolDir "build_report.py") $jsonPath $OutputPath
if ($LASTEXITCODE -ne 0) {
  throw "Excel report build failed. ExitCode=$LASTEXITCODE"
}

Write-Host "Output: $OutputPath"
