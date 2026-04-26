<#
.SYNOPSIS
    OneFinder one-click build and MSI packaging

.DESCRIPTION
    Steps:
      1. msbuild  - Build OneFinder.AddIn (.NET 4.8 COM DLL loaded by OneNote)
      2. dotnet publish  - Publish .NET 8 main app as win-x64 self-contained single-file EXE
      3. wix build  - Package EXE + DLL + registry into MSI

.PREREQUISITES
    - Visual Studio 2022+ (with .NET Framework 4.8 targeting pack)
    - .NET 8 SDK  : winget install Microsoft.DotNet.SDK.8
    - WiX v5      : dotnet tool install --global wix

.EXAMPLE
    .\build.ps1
    .\build.ps1 -Configuration Debug
#>

param(
    [string] $Configuration = "Release"
)

$ErrorActionPreference = "Stop"

$Root         = $PSScriptRoot
$AppProject   = Join-Path $Root "OneFinder\OneFinder.csproj"
$AddinProject = Join-Path $Root "OneFinder.AddIn\OneFinder.AddIn.csproj"
$WxsFile      = Join-Path $Root "installer\Package.wxs"
$PublishDir   = Join-Path $Root "OneFinder\bin\$Configuration\net8.0-windows\win-x64\publish\"
$AddinDir     = Join-Path $Root "OneFinder.AddIn\bin\x64\$Configuration\net48\"
$OutputMsi    = Join-Path $Root "installer\OneFinderSetup.msi"

function Find-MSBuild {
    $vswhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
    if (-not (Test-Path $vswhere)) {
        $vswhere = "D:\Program Files\Microsoft Visual Studio\Installer\vswhere.exe"
    }
    if (Test-Path $vswhere) {
        $vsPath = & $vswhere -latest -products * -requires Microsoft.Component.MSBuild -property installationPath 2>$null
        if ($vsPath) {
            $msb = Join-Path $vsPath "MSBuild\Current\Bin\amd64\MSBuild.exe"
            if (Test-Path $msb) { return $msb }
            $msb = Join-Path $vsPath "MSBuild\Current\Bin\MSBuild.exe"
            if (Test-Path $msb) { return $msb }
        }
    }
    $msb = Get-Command MSBuild.exe -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Source
    if ($msb) { return $msb }
    throw "MSBuild.exe not found. Install Visual Studio 2022 or add MSBuild to PATH."
}

function Require-Command([string]$name) {
    if (-not (Get-Command $name -ErrorAction SilentlyContinue)) {
        Write-Error "Command not found: '$name'. Please install required tools (see script comments)."
        exit 1
    }
}

Require-Command "dotnet"
Require-Command "wix"
$MSBuild = Find-MSBuild
Write-Host "    MSBuild: $MSBuild" -ForegroundColor Gray

# Step 1: Build OneNote addin DLL (.NET Framework 4.8)
Write-Host ""
Write-Host ">>> [1/3] Building OneFinder.AddIn (.NET Framework 4.8)..." -ForegroundColor Cyan

& $MSBuild $AddinProject /p:Configuration=$Configuration /p:Platform=x64 /v:minimal /nologo

if ($LASTEXITCODE -ne 0) { Write-Error "AddIn build failed"; exit 1 }

$AddinDll = Join-Path $AddinDir "OneFinder.AddIn.dll"
if (-not (Test-Path $AddinDll)) {
    Write-Error "AddIn DLL not found: $AddinDll"
    exit 1
}
Write-Host "    DLL: $AddinDll" -ForegroundColor Gray

# Step 2: Publish self-contained single-file EXE
Write-Host ""
Write-Host ">>> [2/3] Publishing OneFinder (self-contained, win-x64, single-file)..." -ForegroundColor Cyan

dotnet publish $AppProject `
    --configuration $Configuration `
    --runtime win-x64 `
    --self-contained true `
    -p:PublishSingleFile=true `
    -p:PublishTrimmed=false `
    -p:DebugType=none `
    -p:DebugSymbols=false

if ($LASTEXITCODE -ne 0) { Write-Error "dotnet publish failed"; exit 1 }

$ExePath = Join-Path $PublishDir "OneFinder.exe"
if (-not (Test-Path $ExePath)) {
    Write-Error "Published EXE not found: $ExePath"
    exit 1
}
Write-Host "    EXE: $ExePath" -ForegroundColor Gray

# Step 3: WiX MSI build
Write-Host ""
Write-Host ">>> [3/3] Building MSI..." -ForegroundColor Cyan

if (-not $PublishDir.EndsWith("\")) { $PublishDir += "\" }
if (-not $AddinDir.EndsWith("\"))   { $AddinDir   += "\" }

wix build $WxsFile `
    -d "PublishDir=$PublishDir" `
    -d "AddinDir=$AddinDir" `
    -arch x64 `
    -out $OutputMsi

if ($LASTEXITCODE -ne 0) { Write-Error "wix build failed"; exit 1 }

Write-Host ""
Write-Host ">>> Done! MSI:" -ForegroundColor Green
Write-Host "    $OutputMsi" -ForegroundColor Yellow