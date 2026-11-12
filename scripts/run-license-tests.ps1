<#
.SYNOPSIS
  Level 1 checks of the licensing module: everything that needs no server and no manual Excel work.
  Nothing is uploaded, the Registry cache is not touched, repo files are not changed.

.DESCRIPTION
  Runs, in this order:
    1. contract-check.ps1            (contract, ASCII, SHA-256 manifest)
    2. build-excel-regression.mjs    (generates TestResults\modLicenseRegression.bas; needs the API repo)
    3. vba-static-check.mjs          (structure and argument counts)
    4. License_RunMockTests              in a separate hidden Excel
    5. License_RunSignedRegressionTests  in the same hidden Excel
  Steps 4 and 5 run on COPIES in out\license-tests\ whose final MsgBox writes cell Z1 instead, because a
  hidden Excel cannot answer a dialog. HTTP, cache and UI are mocked in the regression module.

.EXAMPLE
  .\scripts\run-license-tests.ps1
  .\scripts\run-license-tests.ps1 -ApiRepo D:\Code\digital-license-api -SkipExcel

.NOTES
  Needs Node.js, Excel for Windows and "Trust access to the VBA project object model" (AccessVBOM).
  The Excel PID is written to out\license-tests\excel.pid; a watchdog stops exactly that PID after
  -TimeoutSeconds if Excel hangs (a VBA compile error shows a dialog nobody can answer).
  This file is ASCII only so Windows PowerShell 5.1 reads it the same with or without a BOM.
#>
param(
    [string]$ApiRepo = 'D:\Code\digital-license-api',
    [int]$TimeoutSeconds = 600,
    [switch]$SkipExcel
)
$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
$moduleDir = Join-Path $root 'license-sample\api-v1-one-file'
$work = Join-Path $root 'out\license-tests'
$pidFile = Join-Path $work 'excel.pid'
$resultFile = Join-Path $work 'result.txt'
$XlWbaTWorksheet = -4167
$MsoAutomationSecurityLow = 1
$latin = [Text.Encoding]::GetEncoding(28591)
$results = New-Object System.Collections.Generic.List[object]

function Add-Result([string]$name, [bool]$passed, [string]$detail) {
    $results.Add([pscustomobject]@{ Check = $name; Result = $(if ($passed) { 'PASS' } else { 'FAIL' }); Detail = $detail })
}

# Runs a native command and reports its exit code; the last output line becomes the detail.
function Invoke-Step([string]$name, [scriptblock]$command) {
    $output = & $command | Out-String
    $lastLine = ($output.Trim() -split "`r?`n" | Select-Object -Last 1)
    Add-Result $name ($LASTEXITCODE -eq 0) $lastLine
    return ($LASTEXITCODE -eq 0)
}

# Each replacement must match exactly once, otherwise the source changed and the copy would be wrong.
function Copy-WithReplacements([string]$from, [string]$to, [hashtable[]]$pairs) {
    $code = [IO.File]::ReadAllText($from, $latin)
    foreach ($pair in $pairs) {
        $count = ([regex]::Matches($code, $pair.pattern)).Count
        if ($count -ne 1) { throw "Pattern found $count times in ${from}: $($pair.pattern)" }
        $code = [regex]::Replace($code, $pair.pattern, $pair.with)
    }
    [IO.File]::WriteAllText($to, $code, $latin)
}

New-Item -ItemType Directory -Force $work | Out-Null
if (Test-Path -LiteralPath $resultFile) { Remove-Item -LiteralPath $resultFile }

[void](Invoke-Step 'contract-check' { & powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $moduleDir 'contract-check.ps1') })

$regressionBuilt = $false
if (Test-Path -LiteralPath (Join-Path $ApiRepo 'apps\api\src\modules\licensing\token.service.ts')) {
    $regressionBuilt = Invoke-Step 'build regression module' { & node (Join-Path $moduleDir 'tests\build-excel-regression.mjs') $ApiRepo }
} else {
    Add-Result 'build regression module' $false "API repository not found at $ApiRepo (pass -ApiRepo)"
}
if ($regressionBuilt) {
    [void](Invoke-Step 'static check' { & node (Join-Path $moduleDir 'tests\vba-static-check.mjs') })
} else {
    [void](Invoke-Step 'static check (module only)' { & node (Join-Path $moduleDir 'tests\vba-static-check.mjs') 'modLicenseApiV1.bas' })
}

if ($SkipExcel) {
    Add-Result 'VBA suites in Excel' $false 'skipped (-SkipExcel)'
} else {
    $runs = @()
    $mockCopy = Join-Path $work 'mock.bas'
    Copy-WithReplacements (Join-Path $moduleDir 'modLicenseApiV1.bas') $mockCopy @(
        @{ pattern = 'MsgBox "(\d+) phep thu gia lap dat\.[^\r\n]*'
           with = 'ThisWorkbook.Worksheets(1).Range("Z1").Value = "PASS $1/$1"' },
        @{ pattern = 'MsgBox "Phep thu gia lap that bai: " & CStr\(countPassed\) & "(/\d+)\."[^\r\n]*'
           with = 'ThisWorkbook.Worksheets(1).Range("Z1").Value = "FAIL " & CStr(countPassed) & "$1"' })
    $runs += @{ name = 'License_RunMockTests'; file = $mockCopy; macro = 'modLicenseApiV1.License_RunMockTests' }
    if ($regressionBuilt) {
        $regressionCopy = Join-Path $work 'regression.bas'
        Copy-WithReplacements (Join-Path $moduleDir 'TestResults\modLicenseRegression.bas') $regressionCopy @(
            @{ pattern = 'MsgBox CStr\(passed\) & "(/\d+)[^\r\n]*'
               with = 'ThisWorkbook.Worksheets(1).Range("Z1").Value = "PASS " & CStr(passed) & "$1"' },
            @{ pattern = 'MsgBox "Regression failed after "[^\r\n]*'
               with = 'ThisWorkbook.Worksheets(1).Range("Z1").Value = "FAIL after " & CStr(passed) & ": " & Err.Description' })
        $runs += @{ name = 'License_RunSignedRegressionTests'; file = $regressionCopy
                    macro = 'modLicenseRegression.License_RunSignedRegressionTests' }
    }

    Add-Type -Namespace W -Name U -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(System.IntPtr h, out uint p);'
    $excel = New-Object -ComObject Excel.Application
    [uint32]$excelPid = 0
    [void][W.U]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$excelPid)
    Set-Content -LiteralPath $pidFile -Value $excelPid
    $watchdog = Start-Job -ArgumentList $excelPid, $TimeoutSeconds -ScriptBlock {
        param($id, $seconds) Start-Sleep -Seconds $seconds; Stop-Process -Id $id -Force -ErrorAction SilentlyContinue }
    try {
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.EnableEvents = $false
        $excel.AutomationSecurity = $MsoAutomationSecurityLow
        foreach ($run in $runs) {
            $started = Get-Date
            try {
                $workbook = $excel.Workbooks.Add($XlWbaTWorksheet)
                [void]$workbook.VBProject.VBComponents.Import($run.file)
                $excel.Run($run.macro)
                $cell = [string]$workbook.Worksheets.Item(1).Range('Z1').Value2
                $workbook.Close($false)
                $seconds = [int]((Get-Date) - $started).TotalSeconds
                $allPassed = ($cell -match '^PASS (\d+)/(\d+)$') -and ($Matches[1] -eq $Matches[2])
                Add-Result $run.name $allPassed "$cell (${seconds}s)"
            } catch {
                Add-Result $run.name $false "Excel error (compile error, hang or AccessVBOM off): $($_.Exception.Message)"
            }
        }
    } finally {
        # Cleanup must never stop halfway: Quit can fail if the watchdog already stopped Excel.
        try { $excel.Quit() } catch { Write-Warning "Excel.Quit failed: $($_.Exception.Message)" }
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
        Stop-Job $watchdog -ErrorAction SilentlyContinue
        Remove-Job $watchdog -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $pidFile -ErrorAction SilentlyContinue
    }
}

$table = $results | Format-Table -AutoSize -Wrap | Out-String
Set-Content -LiteralPath $resultFile -Value $table -Encoding UTF8
Write-Host $table
$failed = @($results | Where-Object { $_.Result -ne 'PASS' })
if ($failed.Count -gt 0) {
    Write-Host "$($failed.Count) check(s) did not pass." -ForegroundColor Red
    exit 1
}
Write-Host 'All level 1 checks passed.' -ForegroundColor Green
