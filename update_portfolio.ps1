# ==============================================================================
# SAMUEL ALEMU | AUTOMATIC PORTFOLIO RE-SCAN, ADD & DELETE BUILD SCRIPT
# Run this script whenever you add/delete photos, drawings, or machine entries!
# ==============================================================================

$projectRoot = "c:\Users\admin\Documents\_solidworks\antigravity AI\macro_projects\Part sheet metal DXF generation"
Set-Location $projectRoot

Write-Host "--------------------------------------------------------" -ForegroundColor Cyan
Write-Host " Scanning project folders & Word document list..." -ForegroundColor Cyan
Write-Host "--------------------------------------------------------" -ForegroundColor Cyan

# 1. Run media categorizer script
$categorizeScript = "C:\Users\admin\.gemini\antigravity-ide\brain\7e4b9959-d07c-407e-8a29-7b52abfc3c18\scratch\categorize_media.ps1"
if (Test-Path $categorizeScript) {
    & powershell -ExecutionPolicy Bypass -File $categorizeScript
}

# 2. Run clean dataset generator script (reads Real projects.docx)
$datasetScript = "C:\Users\admin\.gemini\antigravity-ide\brain\7e4b9959-d07c-407e-8a29-7b52abfc3c18\scratch\generate_clean_dataset.ps1"
if (Test-Path $datasetScript) {
    & powershell -ExecutionPolicy Bypass -File $datasetScript
}

# Sync HTML, CSS, JS & assets to deployment folders
Copy-Item "index.html" "portfolio_website\index.html" -Force
Copy-Item "styles.css" "portfolio_website\styles.css" -Force
Copy-Item "app.js" "portfolio_website\app.js" -Force

Copy-Item "index.html" "website\index.html" -Force
Copy-Item "styles.css" "website\styles.css" -Force
Copy-Item "app.js" "website\app.js" -Force

if (Test-Path "assets") {
    Copy-Item "assets\*" "portfolio_website\assets" -Recurse -Force
}

# 6. Copy to User home for easy command line access
Copy-Item "update_portfolio.ps1" "C:\Users\admin\update_portfolio.ps1" -Force

Write-Host ""
Write-Host "========================================================" -ForegroundColor Green
Write-Host " SUCCESS! Portfolio media, text & cards are updated!" -ForegroundColor Green
Write-Host " Refresh your browser at http://localhost:8080/ to view." -ForegroundColor Green
Write-Host "========================================================" -ForegroundColor Green
