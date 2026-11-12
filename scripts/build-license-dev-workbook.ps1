<#
.SYNOPSIS
  Level 2: builds test workbooks straight from the licensing module in this working tree, so a module
  change can be tried against the real server without publishing the VBA package, running the app,
  listing a product or downloading it with another account. Nothing is uploaded.

.DESCRIPTION
  Writes to out\license-dev\ (out\ is gitignored):
    <key>_<version>_dev.xlsm           the test workbook
    <key>_<version>_dev-copy.xlsm      a byte copy of it (two open copies share one cache entry)
    <key>_<version>_dev-offline.xlsm   same module with Cfg_ProxyServer set to a dead local proxy, so every
                                       request fails like a lost connection while the PC stays online
  Each workbook has the real module with the three per-product values filled, the ThisWorkbook events,
  a visible sheet "Huong dan" with a button that calls License_PremiumAllowed (result in the status bar),
  and a protected sheet "Du lieu" saved very hidden.

  The server only sees product_key, version, the device fingerprint and the account that approves the
  code, so the product key and version must be ones that already exist on the store and that the test
  account owns. All three files use the same Registry cache entry Activation_<key>.

.EXAMPLE
  .\scripts\build-license-dev-workbook.ps1 -ProductKey test-01 -Version 1.0.0

.NOTES
  TEST FILES ONLY: they carry a test module (modDevTest). Never give them to a customer.
  Needs Excel for Windows and "Trust access to the VBA project object model" (AccessVBOM).
  Excel runs hidden with macros disabled; its PID is in out\license-dev\excel.pid while the script runs
  and a watchdog stops exactly that PID after 180 seconds if Excel hangs.
  This file is ASCII only so Windows PowerShell 5.1 reads it the same with or without a BOM.
#>
param(
    [Parameter(Mandatory = $true)][string]$ProductKey,
    [Parameter(Mandatory = $true)][string]$Version,
    [string]$OutDir,
    [string]$OfflineProxy = '127.0.0.1:9'
)
$ErrorActionPreference = 'Stop'

if ($ProductKey -notmatch '^[A-Za-z0-9_.-]{1,128}$') { throw "ProductKey '$ProductKey' is not a valid product key." }
if ($Version -notmatch '^[0-9]+\.[0-9]+\.[0-9]+$') { throw "Version '$Version' must look like 1.2.3." }
if ($OfflineProxy -notmatch '^[A-Za-z0-9.-]+:[0-9]{1,5}$') { throw "OfflineProxy '$OfflineProxy' must look like host:port." }

$root = Split-Path -Parent $PSScriptRoot
$sourceDir = Join-Path $root 'license-sample\api-v1-one-file'
$basFile = Join-Path $sourceDir 'modLicenseApiV1.bas'
$eventsFile = Join-Path $sourceDir 'ThisWorkbook-events.txt'
if (-not $OutDir) { $OutDir = Join-Path $root 'out\license-dev' }
# Excel resolves a relative path against its own folder, not this shell's: make it absolute first.
$OutDir = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutDir)
$pidFile = Join-Path $OutDir 'excel.pid'
$baseName = "${ProductKey}_${Version}_dev"
$mainFile = Join-Path $OutDir "$baseName.xlsm"
$copyFile = Join-Path $OutDir "$baseName-copy.xlsm"
$offlineFile = Join-Path $OutDir "$baseName-offline.xlsm"

$guardSeconds = 180
$XlOpenXmlMacroEnabled = 52
$XlWbaTWorksheet = -4167
$XlSheetVisible = -1
$XlSheetVeryHidden = 2
$MsoAutomationSecurityForceDisable = 3
$VbextCtDocument = 100
$ThisWorkbookName = 'ThisWorkbook'
$LicenseModuleName = 'modLicenseApiV1'
$DevModuleName = 'modDevTest'
$GuideSheetName = 'Huong dan'
$ProtectedSheetName = 'Du lieu'
$RequiredHandlers = @('Workbook_BeforeClose', 'Workbook_BeforeSave', 'Workbook_AfterSave')

foreach ($file in $basFile, $eventsFile) {
    if (-not (Test-Path -LiteralPath $file)) { throw "Source file not found: $file" }
}

$devModule = @"
Attribute VB_Name = "$DevModuleName"
Option Explicit

' Test-only helper written by scripts\build-license-dev-workbook.ps1. Never ship this workbook.
Public Sub Dev_CheckAccess()
    ' A refusal in a running session is handled by the licensing module itself: guidance, then close.
    If modLicenseApiV1.License_PremiumAllowed() Then
        Application.StatusBar = "Dev test: con quyen luc " & Format`$(Now, "hh:nn:ss")
    End If
End Sub
"@ -replace "`r?`n", "`r`n"

# Each replacement must match exactly once, otherwise the source changed and the test file would be wrong.
function Get-ConfiguredModule([hashtable[]]$replacements) {
    $code = [IO.File]::ReadAllText($basFile, [Text.Encoding]::ASCII)
    foreach ($item in $replacements) {
        $count = ([regex]::Matches($code, [regex]::Escape($item.from))).Count
        if ($count -ne 1) { throw "Expected one '$($item.from.Trim())' in the module, found $count." }
        $code = $code.Replace($item.from, $item.to)
    }
    return $code
}

function Get-ComponentCode($component) {
    $module = $component.CodeModule
    if ($module.CountOfLines -lt 1) { return '' }
    return $module.Lines(1, $module.CountOfLines)
}

function New-DevWorkbook($excel, [string]$outFile, [string]$moduleCode, [string]$note) {
    $moduleTemp = Join-Path $OutDir 'modLicenseApiV1.bas'
    $devTemp = Join-Path $OutDir 'modDevTest.bas'
    [IO.File]::WriteAllText($moduleTemp, $moduleCode, [Text.Encoding]::ASCII)
    [IO.File]::WriteAllText($devTemp, $devModule, [Text.Encoding]::ASCII)
    $workbook = $excel.Workbooks.Add($XlWbaTWorksheet)
    try {
        try { $components = $workbook.VBProject.VBComponents }
        catch { throw 'Excel refused access to the VBA project. Turn on File > Options > Trust Center > Macro Settings > Trust access to the VBA project object model.' }
        # A new workbook's document module is not always named ThisWorkbook: rename it like a real file.
        foreach ($component in $components) {
            if ($component.Type -eq $VbextCtDocument -and $component.Name -eq $workbook.CodeName) {
                $component.Properties.Item('_CodeName').Value = $ThisWorkbookName
            }
        }
        [void]$components.Import($moduleTemp)
        [void]$components.Import($devTemp)
        $events = (Get-Content -LiteralPath $eventsFile | Where-Object { -not $_.TrimStart().StartsWith("'") }) -join "`r`n"
        $components.Item($ThisWorkbookName).CodeModule.AddFromString($events)

        $guide = $workbook.Worksheets.Item(1)
        $guide.Name = $GuideSheetName
        $guide.Range('A1').Value2 = 'FILE THU NGHIEM module ban quyen. Khong gui cho khach hang.'
        $guide.Range('A2').Value2 = "San pham: $ProductKey   Phien ban: $Version   $note"
        $guide.Range('A3').Value2 = 'Sheet "Du lieu" chi hien khi may nay co quyen.'
        $guide.Range('A4').Value2 = 'Bam nut ben duoi de goi License_PremiumAllowed; con quyen thi thanh trang thai ghi gio kiem tra.'
        $guide.Range('A1').Font.Bold = $true
        $button = $guide.Buttons().Add(12, 90, 300, 30)
        $button.OnAction = "$DevModuleName.Dev_CheckAccess"
        $button.Caption = 'Kiem tra quyen (License_PremiumAllowed)'

        $data = $workbook.Worksheets.Add([Type]::Missing, $guide)
        $data.Name = $ProtectedSheetName
        $data.Range('A1').Value2 = 'Noi dung duoc bao ve. Thay sheet nay nghia la may dang co quyen.'
        $guide.Activate()
        $data.Visible = $XlSheetVeryHidden

        $workbook.SaveAs($outFile, $XlOpenXmlMacroEnabled)
    } finally {
        $workbook.Close($false)
        Remove-Item -LiteralPath $moduleTemp, $devTemp -ErrorAction SilentlyContinue
    }
}

# Re-reads the saved file (macros disabled) and returns the list of problems; empty when it is as intended.
function Get-SavedWorkbookProblems($excel, [string]$file, [string]$expectedProxy) {
    $problems = @()
    $saved = $excel.Workbooks.Open($file, 0, $true)
    try {
        $licenseCode = Get-ComponentCode $saved.VBProject.VBComponents.Item($LicenseModuleName)
        $thisCode = Get-ComponentCode $saved.VBProject.VBComponents.Item($ThisWorkbookName)
        $devCode = Get-ComponentCode $saved.VBProject.VBComponents.Item($DevModuleName)
        $guideVisible = $saved.Worksheets.Item($GuideSheetName).Visible
        $dataVisible = $saved.Worksheets.Item($ProtectedSheetName).Visible
        $buttons = $saved.Worksheets.Item($GuideSheetName).Buttons().Count
    } finally { $saved.Close($false) }
    $expected = [ordered]@{
        Cfg_ProductKey          = "`"$ProductKey`""
        Cfg_ProductVersion      = "`"$Version`""
        Cfg_ProtectedSheetNames = "Array(`"$ProtectedSheetName`")"
        Cfg_ProxyServer         = $expectedProxy
    }
    foreach ($name in $expected.Keys) {
        $found = [regex]::Matches($licenseCode, '(?im)^[ \t]*' + [regex]::Escape($name) + '[ \t]*=[ \t]*(.*?)[ \t]*\r?$')
        if ($found.Count -ne 1) { $problems += "$name has $($found.Count) assignments, expected 1." }
        elseif ($found[0].Groups[1].Value -cne $expected[$name]) { $problems += "$name = $($found[0].Groups[1].Value), expected $($expected[$name])." }
    }
    foreach ($handler in $RequiredHandlers) {
        if ($thisCode -notmatch ('(?i)Sub\s+' + $handler + '\b')) { $problems += "ThisWorkbook lacks $handler." }
    }
    if ($devCode -notmatch 'Public Sub Dev_CheckAccess\(\)') { $problems += 'modDevTest lacks Dev_CheckAccess.' }
    if ($guideVisible -ne $XlSheetVisible) { $problems += "Sheet '$GuideSheetName' is not visible." }
    if ($dataVisible -ne $XlSheetVeryHidden) { $problems += "Sheet '$ProtectedSheetName' is not very hidden." }
    if ($buttons -ne 1) { $problems += "Sheet '$GuideSheetName' has $buttons buttons, expected 1." }
    return $problems
}

$productReplacements = @(
    @{ from = '    Cfg_ProductKey = vbNullString';     to = "    Cfg_ProductKey = `"$ProductKey`"" },
    @{ from = '    Cfg_ProductVersion = vbNullString'; to = "    Cfg_ProductVersion = `"$Version`"" },
    @{ from = '    Cfg_ProtectedSheetNames = Array()'; to = "    Cfg_ProtectedSheetNames = Array(`"$ProtectedSheetName`")" }
)
$onlineCode = Get-ConfiguredModule $productReplacements
$offlineCode = Get-ConfiguredModule ($productReplacements + @(
    @{ from = '    Cfg_ProxyServer = vbNullString'; to = "    Cfg_ProxyServer = `"$OfflineProxy`"" }))

New-Item -ItemType Directory -Force $OutDir | Out-Null
foreach ($file in $mainFile, $copyFile, $offlineFile) {
    if (Test-Path -LiteralPath $file) { Remove-Item -LiteralPath $file }
}

Add-Type -Namespace W -Name U -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(System.IntPtr h, out uint p);'
$excel = New-Object -ComObject Excel.Application
[uint32]$excelPid = 0
[void][W.U]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$excelPid)
Set-Content -LiteralPath $pidFile -Value $excelPid
$watchdog = Start-Job -ArgumentList $excelPid, $guardSeconds -ScriptBlock {
    param($id, $seconds) Start-Sleep -Seconds $seconds; Stop-Process -Id $id -Force -ErrorAction SilentlyContinue }
try {
    Write-Host "Hidden Excel, PID $excelPid (in $pidFile)"
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    # Macros stay disabled for the whole build: importing or reopening must never run the licensing gate.
    $excel.AutomationSecurity = $MsoAutomationSecurityForceDisable

    New-DevWorkbook $excel $mainFile $onlineCode 'Mang: binh thuong.'
    New-DevWorkbook $excel $offlineFile $offlineCode "Mang: GIA LAP MAT MANG (proxy chet $OfflineProxy)."

    $failures = @()
    $failures += @(Get-SavedWorkbookProblems $excel $mainFile 'vbNullString' | ForEach-Object { "$(Split-Path -Leaf $mainFile): $_" })
    $failures += @(Get-SavedWorkbookProblems $excel $offlineFile "`"$OfflineProxy`"" | ForEach-Object { "$(Split-Path -Leaf $offlineFile): $_" })
    if ($failures.Count -gt 0) {
        Remove-Item -LiteralPath $mainFile, $offlineFile -ErrorAction SilentlyContinue
        throw ("The saved test workbooks are not as intended:`n - " + ($failures -join "`n - "))
    }
    Copy-Item -LiteralPath $mainFile -Destination $copyFile
} finally {
    # Cleanup must never stop halfway: Quit can fail if the watchdog already stopped Excel.
    try { $excel.Quit() } catch { Write-Warning "Excel.Quit failed: $($_.Exception.Message)" }
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
    Stop-Job $watchdog -ErrorAction SilentlyContinue
    Remove-Job $watchdog -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $pidFile -ErrorAction SilentlyContinue
}

Write-Host "Built for product '$ProductKey' version $Version (cache entry Activation_$ProductKey):"
Write-Host "  $mainFile"
Write-Host "  $copyFile"
Write-Host "  $offlineFile   (requests fail like a lost connection)"
Write-Host 'Test files only: never give them to a customer.'
