$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location -LiteralPath $ProjectRoot

# Create venv if missing
if (-not (Test-Path -LiteralPath ".venv")) {
    Write-Host "Creating virtual environment..." -ForegroundColor Green
    python -m venv .venv
}

# Activate
.\.venv\Scripts\Activate.ps1

# Install deps
if (Test-Path -LiteralPath "requirements.txt") {
    Write-Host "Installing dependencies..." -ForegroundColor Green
    pip install -r requirements.txt
}

# Run
Write-Host "Starting parser..." -ForegroundColor Green
python src/doors_parser.py --targets data/target_lines.csv --sites config/sites.yaml --out output
