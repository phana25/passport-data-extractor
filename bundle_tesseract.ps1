$ErrorActionPreference = "Stop"

$sourceCandidates = @(
    "C:\Program Files\Tesseract-OCR",
    "C:\Program Files (x86)\Tesseract-OCR"
)

$sourceDir = $null
foreach ($candidate in $sourceCandidates) {
    if (Test-Path "$candidate\tesseract.exe") {
        $sourceDir = $candidate
        break
    }
}

if (-not $sourceDir) {
    throw "Tesseract install not found. Install Tesseract first, then rerun this script."
}

$targetDir = Join-Path $PSScriptRoot "vendor\tesseract"
New-Item -ItemType Directory -Path $targetDir -Force | Out-Null

Copy-Item "$sourceDir\tesseract.exe" $targetDir -Force
Get-ChildItem $sourceDir -Filter "*.dll" | ForEach-Object {
    Copy-Item $_.FullName $targetDir -Force
}

$sourceTessdata = Join-Path $sourceDir "tessdata"
$targetTessdata = Join-Path $targetDir "tessdata"
New-Item -ItemType Directory -Path $targetTessdata -Force | Out-Null

# Keep bundle small: include only required language/model files.
$filesToCopy = @("eng.traineddata", "osd.traineddata", "*.config")
foreach ($pattern in $filesToCopy) {
    Get-ChildItem $sourceTessdata -Filter $pattern -ErrorAction SilentlyContinue | ForEach-Object {
        Copy-Item $_.FullName $targetTessdata -Force
    }
}

Write-Host "Bundled Tesseract to: $targetDir"
Write-Host "Now rebuild using your .spec file."
