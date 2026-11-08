<#
.SYNOPSIS
  Dựng các file .xlsm thử tay cho LicenseGateSmokeTest.md (ca đóng file và ca hai Auto_Open).
  Không đụng file thật, không gọi mạng, không tải gì lên.

.DESCRIPTION
  Ghi vào out\manual-test\ (out\ đã gitignore):
  - close-test.xlsm: sản phẩm giả "close-test" 1.0.0. Module bản quyền thật, nhưng Auto_Open đổi tên thành
    Auto_Open_TestDisabled để mở file không gọi mạng và không tự đóng. ThisWorkbook có BeforeClose của sản phẩm
    (ghi "BeforeClose da chay" vào Huong dan!B1, dòng bản quyền đặt CUỐI), BeforeSave, AfterSave.
    Sheet "Du lieu" được liệt kê trong Cfg_ProtectedSheetNames và đang HIỆN.
  - auto-open-license-first.xlsm, auto-open-other-first.xlsm: hai module giả đều có Auto_Open, ghi "license" hoặc
    "other" vào cột A, chỉ khác thứ tự tạo module.

.EXAMPLE
  .\tests\manual\build-manual-test-files.ps1

.NOTES
  Cần Excel Windows và AccessVBOM đang bật. Excel chạy ẩn; PID ghi vào excel.pid cạnh các file. Một job canh
  dừng đúng PID đó sau 180 giây nếu Excel treo. Lưu UTF-8 có BOM cho Windows PowerShell 5.1.
#>
param(
    [string]$OutDir
)
$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$basFile = Join-Path $root 'license-sample\api-v1-one-file\modLicenseApiV1.bas'
if (-not $OutDir) { $OutDir = Join-Path $root 'out\manual-test' }
# Excel resolves a relative path against its own folder, not this shell's: make it absolute first.
$OutDir = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutDir)
$guardSeconds = 180
$XlOpenXmlMacroEnabled = 52
$XlWbaTWorksheet = -4167
$MsoAutomationSecurityForceDisable = 3
$VbextCtStdModule = 1
$VbextCtDocument = 100

# Mỗi thay thế phải khớp đúng một lần, nếu không source đã đổi và file thử sẽ sai.
$Replacements = @(
    @{ from = '    Cfg_ProductKey = vbNullString';           to = '    Cfg_ProductKey = "close-test"' },
    @{ from = '    Cfg_ProductVersion = vbNullString';       to = '    Cfg_ProductVersion = "1.0.0"' },
    @{ from = '    Cfg_ProtectedSheetNames = Array()';       to = '    Cfg_ProtectedSheetNames = Array("Du lieu")' },
    @{ from = 'Public Sub Auto_Open()';                      to = 'Public Sub Auto_Open_TestDisabled()' }
)

$closeTestEvents = @'
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ThisWorkbook.Worksheets("Huong dan").Range("B1").Value = "BeforeClose da chay"
'=== tool-vba BEGIN license hook (auto-managed, do not edit) ===
    If Not Cancel And (ThisWorkbook.Saved) Then modLicenseApiV1.License_OnClose
'=== tool-vba END license hook ===
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    If Not modLicenseApiV1.License_BeforeSave() Then Cancel = True: Exit Sub
End Sub

Private Sub Workbook_AfterSave(ByVal Success As Boolean)
    modLicenseApiV1.License_AfterSave
End Sub
'@

# The VBE expects CR/LF; this file's own line endings depend on the checkout (core.autocrlf).
$closeTestEvents = $closeTestEvents -replace "`r?`n", "`r`n"

function Get-TestLicenseCode {
    $code = [IO.File]::ReadAllText($basFile, [Text.Encoding]::ASCII)
    foreach ($item in $Replacements) {
        $count = ([regex]::Matches($code, [regex]::Escape($item.from))).Count
        if ($count -ne 1) { throw "Source đã đổi: '$($item.from)' khớp $count lần, cần đúng 1." }
        $code = $code.Replace($item.from, $item.to)
    }
    return $code
}

function Get-AutoOpenStubCode([string]$tag) {
    return "Public Sub Auto_Open()`r`n    With ThisWorkbook.Worksheets(1)`r`n" +
        "        .Cells(.Rows.Count, 1).End(xlUp).Offset(1, 0).Value = `"$tag`"`r`n    End With`r`nEnd Sub`r`n"
}

function Rename-ThisWorkbook($workbook) {
    foreach ($component in $workbook.VBProject.VBComponents) {
        if ($component.Type -eq $VbextCtDocument -and $component.Name -eq $workbook.CodeName) {
            $component.Properties.Item('_CodeName').Value = 'ThisWorkbook'
        }
    }
}

function New-CloseTestFile($excel, [string]$path) {
    $tempBas = Join-Path $OutDir 'close-test-module.bas'
    [IO.File]::WriteAllText($tempBas, (Get-TestLicenseCode), [Text.Encoding]::ASCII)
    try {
        $workbook = $excel.Workbooks.Add($XlWbaTWorksheet)
        $guide = $workbook.Worksheets.Item(1); $guide.Name = 'Huong dan'
        $guide.Range('A1').Value2 = 'File thu tay (close-test). Xem tests/manual/LicenseGateSmokeTest.md.'
        $data = $workbook.Worksheets.Add([Type]::Missing, $guide); $data.Name = 'Du lieu'
        $data.Range('A1').Value2 = 'Du lieu bao ve'
        Rename-ThisWorkbook $workbook
        [void]$workbook.VBProject.VBComponents.Import($tempBas)
        $workbook.VBProject.VBComponents.Item('ThisWorkbook').CodeModule.AddFromString($closeTestEvents)
        $guide.Activate()
        $workbook.SaveAs($path, $XlOpenXmlMacroEnabled); $workbook.Close($false)
    } finally { Remove-Item -LiteralPath $tempBas -ErrorAction SilentlyContinue }
}

function New-AutoOpenPairFile($excel, [string]$path, [string[]]$moduleOrder) {
    $workbook = $excel.Workbooks.Add($XlWbaTWorksheet)
    $workbook.Worksheets.Item(1).Range('A1').Value2 = 'log'
    foreach ($name in $moduleOrder) {
        $component = $workbook.VBProject.VBComponents.Add($VbextCtStdModule)
        $component.Name = $name
        $tag = if ($name -eq 'modLicenseApiV1') { 'license' } else { 'other' }
        $component.CodeModule.AddFromString((Get-AutoOpenStubCode $tag))
    }
    $workbook.SaveAs($path, $XlOpenXmlMacroEnabled); $workbook.Close($false)
}

New-Item -ItemType Directory -Force $OutDir | Out-Null
$files = [ordered]@{
    close        = Join-Path $OutDir 'close-test.xlsm'
    licenseFirst = Join-Path $OutDir 'auto-open-license-first.xlsm'
    otherFirst   = Join-Path $OutDir 'auto-open-other-first.xlsm'
}
foreach ($file in $files.Values) { if (Test-Path -LiteralPath $file) { Remove-Item -LiteralPath $file } }
[void](Get-TestLicenseCode)   # kiểm source trước khi mở Excel

Add-Type -Namespace W -Name U -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(System.IntPtr h, out uint p);'
$excel = New-Object -ComObject Excel.Application
$pidFile = Join-Path $OutDir 'excel.pid'
$guard = $null
try {
    [uint32]$excelPid = 0
    [void][W.U]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$excelPid)
    Set-Content -LiteralPath $pidFile -Value $excelPid
    # The guard stops the PID only while it is still the same Excel process (name and start time).
    $startTicks = (Get-Process -Id $excelPid).StartTime.Ticks
    $guard = Start-Job -ScriptBlock {
        param($id, $ticks, $sec)
        Start-Sleep -Seconds $sec
        $process = Get-Process -Id $id -ErrorAction SilentlyContinue
        if ($process -and $process.ProcessName -eq 'EXCEL' -and $process.StartTime.Ticks -eq $ticks) { Stop-Process -Id $id -Force -ErrorAction SilentlyContinue }
    } -ArgumentList $excelPid, $startTicks, $guardSeconds
    Write-Host "Excel ẩn, PID $excelPid (ghi ở $pidFile; job canh dừng đúng PID này sau $guardSeconds giây)"
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = $MsoAutomationSecurityForceDisable
    New-CloseTestFile $excel $files.close
    New-AutoOpenPairFile $excel $files.licenseFirst @('modLicenseApiV1', 'modOther')
    New-AutoOpenPairFile $excel $files.otherFirst @('modOther', 'modLicenseApiV1')
    foreach ($file in $files.Values) { Write-Host "Đã tạo: $file" }
} finally {
    # Cleanup must never stop halfway: Quit can fail if Excel is already gone (e.g. the guard stopped it).
    try { $excel.Quit() } catch { Write-Warning "Excel.Quit không chạy được: $($_.Exception.Message)" }
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
    if ($guard) {
        Stop-Job $guard -ErrorAction SilentlyContinue
        Remove-Job $guard -Force -ErrorAction SilentlyContinue
    }
    # Drop the PID file once Excel is gone, so it can never point at a reused PID.
    Remove-Item -LiteralPath $pidFile -ErrorAction SilentlyContinue
}
