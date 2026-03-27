param(
    [string]$Configuration = "Release",
    [string]$Version = "1.0.2"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$projectPath = Join-Path $repoRoot "InstantTranslateWin.App\InstantTranslateWin.App.csproj"
$toolsDir = Join-Path $repoRoot ".tools"
$wixPath = Join-Path $toolsDir "wix.exe"

$publishX64Dir = Join-Path $repoRoot "artifacts\publish\win-x64"
$publishX86Dir = Join-Path $repoRoot "artifacts\publish\win-x86"
$setupDir = Join-Path $repoRoot "artifacts\setup"

$x64Wxs = Join-Path $PSScriptRoot "InstantTranslateWin.Setup.x64.wxs"
$x86Wxs = Join-Path $PSScriptRoot "InstantTranslateWin.Setup.x86.wxs"
$x64BundleWxs = Join-Path $PSScriptRoot "InstantTranslateWin.Bundle.x64.wxs"
$x86BundleWxs = Join-Path $PSScriptRoot "InstantTranslateWin.Bundle.x86.wxs"

$versionSuffix = $Version
$x64MsiPath = Join-Path $setupDir "InstantTranslateWin-Setup-$versionSuffix-x64.msi"
$x86MsiPath = Join-Path $setupDir "InstantTranslateWin-Setup-$versionSuffix-x86.msi"
$x64ExePath = Join-Path $setupDir "InstantTranslateWin-Setup-$versionSuffix-x64.exe"
$x86ExePath = Join-Path $setupDir "InstantTranslateWin-Setup-$versionSuffix-x86.exe"

if ($Version -notmatch '^\d+\.\d+\.\d+$')
{
    throw "Version phải theo định dạng MAJOR.MINOR.PATCH, ví dụ 1.0.0"
}

$bundleVersion = "$Version.0"

New-Item -ItemType Directory -Force $toolsDir | Out-Null
New-Item -ItemType Directory -Force $publishX64Dir | Out-Null
New-Item -ItemType Directory -Force $publishX86Dir | Out-Null
New-Item -ItemType Directory -Force $setupDir | Out-Null

Get-ChildItem -Path $publishX64Dir -Force -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
Get-ChildItem -Path $publishX86Dir -Force -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue

if ($Configuration -ieq "Release")
{
    Get-ChildItem -Path $setupDir -Force -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
}
else
{
    Get-ChildItem -Path $setupDir -Filter "InstantTranslateWin-Setup-*" -Force -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
}

if (-not (Test-Path $wixPath))
{
    dotnet tool install --tool-path $toolsDir wix
}

& $wixPath extension add WixToolset.BootstrapperApplications.wixext/6.0.2

dotnet publish $projectPath -c $Configuration -r win-x64 --self-contained true -p:Version=$Version -o $publishX64Dir
dotnet publish $projectPath -c $Configuration -r win-x86 --self-contained true -p:Version=$Version -o $publishX86Dir

& $wixPath build $x64Wxs -d PublishDir=$publishX64Dir -d ProductVersion=$Version -arch x64 -out $x64MsiPath
& $wixPath build $x86Wxs -d PublishDir=$publishX86Dir -d ProductVersion=$Version -arch x86 -out $x86MsiPath

& $wixPath build $x64BundleWxs -d MsiPath=$x64MsiPath -d BundleVersion=$bundleVersion -d DisplayVersion=$Version -ext WixToolset.BootstrapperApplications.wixext -out $x64ExePath
& $wixPath build $x86BundleWxs -d MsiPath=$x86MsiPath -d BundleVersion=$bundleVersion -d DisplayVersion=$Version -ext WixToolset.BootstrapperApplications.wixext -out $x86ExePath

Get-ChildItem $setupDir -Filter "*.exe" | Select-Object FullName, Length, LastWriteTime
