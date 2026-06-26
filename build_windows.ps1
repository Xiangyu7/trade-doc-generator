param(
    [switch]$SkipLibreOfficeBundleCheck,
    [switch]$SkipInstaller
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $Root

$ManifestPath = Join-Path $Root "packaging_manifest.json"
$Manifest = Get-Content -Raw -Encoding UTF8 -LiteralPath $ManifestPath | ConvertFrom-Json
$RequiredTemplates = @($Manifest.required_templates)

Write-Host "Checking business templates..."
$MissingTemplates = @()
foreach ($Name in $RequiredTemplates) {
    $Path = Join-Path $Root "resources\templates\$Name"
    if (-not (Test-Path -LiteralPath $Path)) {
        $MissingTemplates += $Name
    }
}
if ($MissingTemplates.Count -gt 0) {
    Write-Host "Missing templates under resources\templates:" -ForegroundColor Red
    $MissingTemplates | ForEach-Object { Write-Host "  - $_" -ForegroundColor Red }
    throw "Please copy the required business templates before building the installer."
}

if (-not $SkipLibreOfficeBundleCheck) {
    $PortableSoffice = Join-Path $Root "resources\LibreOfficePortable\App\libreoffice\program\soffice.exe"
    $BundledSoffice = Join-Path $Root "resources\libreoffice\program\soffice.exe"
    $LibreOfficeMsi = Join-Path $Root "resources\installers\LibreOffice.msi"
    if (-not ((Test-Path -LiteralPath $PortableSoffice) -or (Test-Path -LiteralPath $BundledSoffice) -or (Test-Path -LiteralPath $LibreOfficeMsi))) {
        Write-Host "Missing bundled LibreOffice:" -ForegroundColor Red
        Write-Host "  resources\LibreOfficePortable\App\libreoffice\program\soffice.exe" -ForegroundColor Red
        Write-Host "  resources\libreoffice\program\soffice.exe" -ForegroundColor Red
        Write-Host "  resources\installers\LibreOffice.msi" -ForegroundColor Red
        throw "Bundle LibreOffice, or rebuild with -SkipLibreOfficeBundleCheck if target PCs already have LibreOffice."
    }
}

Write-Host "Installing runtime and build dependencies..."
python -m pip install -r requirements.txt
python -m pip install -r requirements-build.txt

Write-Host "Building application with PyInstaller..."
python -m PyInstaller --clean --noconfirm trade_doc_generator.spec

if ($SkipInstaller) {
    Write-Host "Skipped installer. App folder: dist\TradeDocGenerator" -ForegroundColor Yellow
    exit 0
}

$IsccCommand = Get-Command ISCC.exe -ErrorAction SilentlyContinue
$IsccPath = if ($IsccCommand) { $IsccCommand.Source } else { $null }
if (-not $IsccPath) {
    $Candidates = @(
        "$env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe",
        "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe",
        "$env:ProgramFiles\Inno Setup 6\ISCC.exe"
    )
    foreach ($Candidate in $Candidates) {
        if ($Candidate -and (Test-Path -LiteralPath $Candidate)) {
            $IsccPath = $Candidate
            break
        }
    }
}
if (-not $IsccPath) {
    throw "Inno Setup 6 was not found. Install it, then rerun this script."
}

Write-Host "Building installer with Inno Setup..."
& $IsccPath "installer\trade-doc-generator.iss"
Write-Host "Done. Installer output: dist-installer" -ForegroundColor Green
