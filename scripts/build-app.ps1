<#
.SYNOPSIS
  Builds tool-vba and packs it with Velopack into out\release\upload, ready to upload in the web Admin
  (Tải công cụ → Phát hành tool-vba → Phát hành ứng dụng). Nothing is uploaded and no password is asked here.

.EXAMPLE
  .\scripts\build-app.ps1 -Version 1.0.1

.NOTES
  Requires: .NET 10 SDK; vpk (dotnet tool install -g vpk --version 1.2.161).
  Reads ToolVbaApiBaseUrl and ToolVbaReleasePublicKey from src\ToolVba.App\release.props.
  Saved as UTF-8 with BOM so Windows PowerShell 5.1 reads the Vietnamese text correctly.
#>
param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^\d+\.\d+\.\d+$')]
    [string]$Version
)
$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
$props = [xml](Get-Content (Join-Path $root 'src\ToolVba.App\release.props') -Raw -Encoding UTF8)
$api = $props.SelectSingleNode('//ToolVbaApiBaseUrl').InnerText.Trim().TrimEnd('/')
$key = $props.SelectSingleNode('//ToolVbaReleasePublicKey').InnerText.Trim()
if (-not $api -or -not $key) { throw 'Hãy điền ToolVbaApiBaseUrl và ToolVbaReleasePublicKey trong src\ToolVba.App\release.props trước.' }
# The update feed travels over this URL: https only.
if (-not $api.StartsWith('https://')) { throw 'ToolVbaApiBaseUrl phải dùng https://.' }

# Runtime the setup installs if missing; keep in step with <TargetFramework> of ToolVba.App.
$DotnetRuntime = 'net10-x64-desktop'

$work = Join-Path $root 'out\release'
$publishDir = Join-Path $work 'publish'
$releasesDir = Join-Path $work 'releases'
$uploadDir = Join-Path $work 'upload'
foreach ($dir in @($publishDir, $uploadDir)) { if (Test-Path $dir) { Remove-Item -Recurse -Force $dir } }
New-Item -ItemType Directory -Force $releasesDir, $uploadDir | Out-Null

Write-Host "1/3 Build tool-vba $Version"
# Framework-dependent: the bundled .NET runtime made each file ~70 MB, over Cloudflare's 100 MB request
# limit for one upload. The Velopack setup installs the .NET Desktop Runtime when a machine lacks it.
dotnet publish (Join-Path $root 'src\ToolVba.App\ToolVba.App.csproj') -c Release -r win-x64 --self-contained false "-p:Version=$Version" -o $publishDir
if ($LASTEXITCODE -ne 0) { throw 'dotnet publish thất bại.' }

Write-Host '2/3 Tải bản trước từ máy chủ (để tạo gói cập nhật nhỏ)'
vpk download http --url "$api/api/v1/tool/app/" --outputDir $releasesDir
if ($LASTEXITCODE -ne 0) { Write-Host '   Chưa có bản trước trên máy chủ (lần phát hành đầu) - tiếp tục.' }

Write-Host '3/3 Đóng gói Velopack'
vpk pack --packId ToolVba --packVersion $Version --packDir $publishDir --mainExe tool-vba.exe --packTitle 'Dong goi VBA' --framework $DotnetRuntime --runtime win-x64 --outputDir $releasesDir
if ($LASTEXITCODE -ne 0) { throw 'vpk pack thất bại.' }

$names = @('releases.win.json', "ToolVba-$Version-full.nupkg", "ToolVba-$Version-delta.nupkg", 'ToolVba-win-Setup.exe')
foreach ($name in $names) {
    $file = Join-Path $releasesDir $name
    if (Test-Path $file) { Copy-Item $file $uploadDir }
}
if (-not (Test-Path (Join-Path $uploadDir 'releases.win.json')) -or -not (Test-Path (Join-Path $uploadDir "ToolVba-$Version-full.nupkg"))) {
    throw 'Không tìm thấy file phát hành do vpk tạo ra.'
}

Write-Host ''
Write-Host "Xong. Tải tất cả file trong thư mục sau lên Admin web (Tải công cụ → Phát hành tool-vba → Phát hành ứng dụng):"
Write-Host "  $uploadDir"
Get-ChildItem $uploadDir | ForEach-Object { Write-Host ("  - {0}  ({1:N1} MB)" -f $_.Name, ($_.Length / 1MB)) }
Start-Process explorer.exe $uploadDir
