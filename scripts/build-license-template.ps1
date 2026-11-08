<#
.SYNOPSIS
  Dựng file .xlsm mẫu chứa module bản quyền, để tải lên Admin web (Phát hành tool-vba → Xuất bản code VBA).
  Không tải gì lên, không hỏi mật khẩu.

.DESCRIPTION
  Nguồn: license-sample\api-v1-one-file\modLicenseApiV1.bas và ThisWorkbook-events.txt (bỏ dòng chú thích).
  Ba giá trị theo sản phẩm PHẢI để trống trong mẫu: Cfg_ProductKey và Cfg_ProductVersion là vbNullString,
  Cfg_ProtectedSheetNames là Array(). tool-vba điền chúng khi bấm Tạo. Sau khi lưu, script mở lại file
  và kiểm; sai thì báo lỗi, thoát mã 1.

.EXAMPLE
  .\scripts\build-license-template.ps1
  .\scripts\build-license-template.ps1 -OutDir D:\Temp\mau

.NOTES
  Cần Excel Windows và quyền "Trust access to the VBA project object model" (AccessVBOM) đang bật.
  Excel chạy ẩn. PID ghi vào excel.pid trong thư mục đầu ra (xóa khi script kết thúc); nếu Excel treo
  (lỗi biên dịch VBA, hộp thoại), từ cửa sổ khác chỉ dừng đúng PID đó: Stop-Process -Id (Get-Content <OutDir>\excel.pid).
  Lưu dạng UTF-8 có BOM để Windows PowerShell 5.1 đọc đúng tiếng Việt.
#>
param(
    [string]$OutDir
)
$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
$sourceDir = Join-Path $root 'license-sample\api-v1-one-file'
$basFile = Join-Path $sourceDir 'modLicenseApiV1.bas'
$eventsFile = Join-Path $sourceDir 'ThisWorkbook-events.txt'
if (-not $OutDir) { $OutDir = Join-Path $root 'out\license-template' }
# Excel resolves a relative path against its own folder, not this shell's: make it absolute first.
$OutDir = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutDir)
$outFile = Join-Path $OutDir 'license-template.xlsm'
$pidFile = Join-Path $OutDir 'excel.pid'

$XlOpenXmlMacroEnabled = 52
$XlWbaTWorksheet = -4167
$MsoAutomationSecurityForceDisable = 3
$VbextCtDocument = 100
$ThisWorkbookName = 'ThisWorkbook'
$LicenseModuleName = 'modLicenseApiV1'
$RequiredHandlers = @('Workbook_BeforeClose', 'Workbook_BeforeSave', 'Workbook_AfterSave')
# Ba giá trị theo sản phẩm và dạng trống bắt buộc của chúng.
$BlankConfig = [ordered]@{
    Cfg_ProductKey          = 'vbNullString'
    Cfg_ProductVersion      = 'vbNullString'
    Cfg_ProtectedSheetNames = 'Array()'
}

foreach ($file in $basFile, $eventsFile) {
    if (-not (Test-Path -LiteralPath $file)) { throw "Không thấy file nguồn: $file" }
}

# Trả về danh sách lỗi (rỗng nếu ba dòng Cfg đều trống đúng dạng).
function Get-BlankConfigProblems([string]$code) {
    $problems = @()
    foreach ($name in $BlankConfig.Keys) {
        $pattern = '(?im)^[ \t]*' + [regex]::Escape($name) + '[ \t]*=[ \t]*(.*?)[ \t]*\r?$'
        $found = [regex]::Matches($code, $pattern)
        if ($found.Count -ne 1) {
            $problems += "$name có $($found.Count) dòng gán, cần đúng 1."
        } elseif ($found[0].Groups[1].Value -ine $BlankConfig[$name]) {
            $problems += "$name = $($found[0].Groups[1].Value), cần $($BlankConfig[$name])."
        }
    }
    return $problems
}

function Get-ComponentCode($component) {
    $module = $component.CodeModule
    if ($module.CountOfLines -lt 1) { return '' }
    return $module.Lines(1, $module.CountOfLines)
}

New-Item -ItemType Directory -Force $OutDir | Out-Null
if (Test-Path -LiteralPath $outFile) { Remove-Item -LiteralPath $outFile }

Add-Type -Namespace W -Name U -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(System.IntPtr h, out uint p);'
$excel = New-Object -ComObject Excel.Application
try {
    [uint32]$excelPid = 0
    [void][W.U]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$excelPid)
    Set-Content -LiteralPath $pidFile -Value $excelPid
    Write-Host "Excel ẩn, PID $excelPid (ghi ở $pidFile)"
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = $MsoAutomationSecurityForceDisable
    $workbook = $excel.Workbooks.Add($XlWbaTWorksheet)
    try { $components = $workbook.VBProject.VBComponents }
    catch { throw 'Excel không cho truy cập VBA project. Bật File → Options → Trust Center → Macro Settings → Trust access to the VBA project object model.' }

    # Module workbook mới có tên dạng Sheet/This...: đổi thành ThisWorkbook như một file thật.
    foreach ($component in $components) {
        if ($component.Type -eq $VbextCtDocument -and $component.Name -eq $workbook.CodeName) {
            $component.Properties.Item('_CodeName').Value = $ThisWorkbookName
        }
    }
    [void]$components.Import($basFile)
    $events = (Get-Content -LiteralPath $eventsFile | Where-Object { -not $_.TrimStart().StartsWith("'") }) -join "`r`n"
    $components.Item($ThisWorkbookName).CodeModule.AddFromString($events)

    $workbook.SaveAs($outFile, $XlOpenXmlMacroEnabled)
    $workbook.Close($false)

    # Kiểm lại từ file đã lưu, không tin vào bản trong bộ nhớ.
    $saved = $excel.Workbooks.Open($outFile, 0, $true)
    try {
        $names = @($saved.VBProject.VBComponents | ForEach-Object { $_.Name })
        $licenseCode = Get-ComponentCode $saved.VBProject.VBComponents.Item($LicenseModuleName)
        $thisCode = Get-ComponentCode $saved.VBProject.VBComponents.Item($ThisWorkbookName)
    } finally { $saved.Close($false) }

    $failures = @(Get-BlankConfigProblems $licenseCode)
    foreach ($handler in $RequiredHandlers) {
        if ($thisCode -notmatch ('(?i)Sub\s+' + $handler + '\b')) { $failures += "ThisWorkbook thiếu $handler." }
    }
    if ($failures.Count -gt 0) {
        Remove-Item -LiteralPath $outFile
        throw ("Kiểm file mẫu thất bại:`n - " + ($failures -join "`n - "))
    }
    Write-Host "Module trong file: $($names -join ', ')"
    Write-Host "Ba dòng Cfg để trống: $($BlankConfig.Keys -join ', ')"
    Write-Host "Xong: $outFile"
} finally {
    # Cleanup must never stop halfway: Quit can fail if Excel is already gone.
    try { $excel.Quit() } catch { Write-Warning "Excel.Quit không chạy được: $($_.Exception.Message)" }
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
    # Drop the PID file once Excel is gone, so it can never point at a reused PID.
    Remove-Item -LiteralPath $pidFile -ErrorAction SilentlyContinue
}
