$ErrorActionPreference = "Stop"

$distAppExe = Join-Path $PSScriptRoot "dist\Passport-Data-Extractor\Passport-Data-Extractor.exe"
if (-not (Test-Path $distAppExe)) {
    throw "Build app first: pyinstaller --noconfirm `"Passport-Data-Extractor.spec`""
}

$vcRedist = Join-Path $PSScriptRoot "third_party\vc_redist.x64.exe"
if (-not (Test-Path $vcRedist)) {
    throw "Missing $vcRedist. Download vc_redist.x64.exe and place it there."
}

$isccCandidates = @(
    "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe",
    "${env:ProgramFiles}\Inno Setup 6\ISCC.exe",
    "${env:LOCALAPPDATA}\Programs\Inno Setup 6\ISCC.exe"
)

$iscc = $null
foreach ($candidate in $isccCandidates) {
    if (Test-Path $candidate) {
        $iscc = $candidate
        break
    }
}

if (-not $iscc) {
    throw "Inno Setup not found. Install Inno Setup 6 first."
}

& $iscc (Join-Path $PSScriptRoot "installer.iss")
Write-Host "Installer created in installer_output\"
