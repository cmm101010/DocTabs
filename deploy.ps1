# 把刚编译好的 DocTabs.exe 覆盖到指定安装目录。
# Office 会映射加载项文件，覆盖前请关闭所有 Word / Excel 窗口。
#
# 用法：
#   .\deploy.ps1 -InstallDir "C:\Tools\DocTabs"

param(
    [Parameter(Mandatory = $true)]
    [string]$InstallDir
)

$ErrorActionPreference = 'Stop'

$src = Join-Path $PSScriptRoot 'stage\DocTabs.exe'
$dst = Join-Path $InstallDir 'DocTabs.exe'

if (-not (Test-Path -LiteralPath $src)) { throw "build first: $src" }
if (-not (Test-Path -LiteralPath $InstallDir)) { throw "install dir not found: $InstallDir" }
if (-not (Test-Path -LiteralPath $dst)) { throw "DocTabs.exe not found in: $InstallDir" }

try {
    $fs = [IO.File]::Open($dst, 'Open', 'ReadWrite', 'None')
    $fs.Close()
}
catch {
    Write-Host 'DocTabs.exe 正被占用：先关掉所有 Word / Excel 窗口再运行本脚本。' -ForegroundColor Yellow
    exit 1
}

$old = (Get-Item -LiteralPath $dst).VersionInfo.FileVersion
$bak = Join-Path $InstallDir ("DocTabs-{0}.exe.bak" -f ($old -replace '\.0$', ''))
Copy-Item -LiteralPath $dst -Destination $bak -Force
Copy-Item -LiteralPath $src -Destination $dst -Force

$new = (Get-Item -LiteralPath $dst).VersionInfo.FileVersion
Write-Output "backup   $bak"
Write-Output "installed $dst  ($old -> $new)"
Write-Output '重新打开 Word 生效。'
