param(
    [string]$Model = "qwen3-vl:8b",
    [string]$Python = "python",
    [string]$ActivationServer,
    [string]$ActivationPublicKey
)

$ErrorActionPreference = "Stop"

if ([string]::IsNullOrWhiteSpace($ActivationServer)) {
    throw "ActivationServer is required. Use the HTTPS URL of your activation server."
}
if ([string]::IsNullOrWhiteSpace($ActivationPublicKey)) {
    throw "ActivationPublicKey is required. Get it from /v1/public-key on your activation server."
}

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $projectRoot

$ollamaCommand = Get-Command ollama.exe -ErrorAction SilentlyContinue
if (-not $ollamaCommand) {
    throw "ollama.exe was not found on PATH. Install Ollama for Windows first."
}
$ollamaPath = $ollamaCommand.Source
$ollamaInstall = Split-Path -Parent $ollamaPath

$modelsSource = $env:OLLAMA_MODELS
if ([string]::IsNullOrWhiteSpace($modelsSource)) {
    $modelsSource = Join-Path $env:USERPROFILE ".ollama\models"
}
New-Item $modelsSource -ItemType Directory -Force | Out-Null

Write-Host "Installing Python dependencies..."
& $Python -m pip install -r requirements.txt
if ($LASTEXITCODE -ne 0) {
    throw "Python dependency installation failed."
}

Write-Host "Downloading $Model if necessary..."
$env:OLLAMA_MODELS = $modelsSource
$env:OLLAMA_HOST = "127.0.0.1:11436"
$buildServer = $null
try {
    $buildServer = Start-Process -FilePath $ollamaPath -ArgumentList "serve" `
        -WorkingDirectory $ollamaInstall -WindowStyle Hidden -PassThru
    $ready = $false
    for ($attempt = 0; $attempt -lt 60; $attempt++) {
        try {
            Invoke-WebRequest -Uri "http://127.0.0.1:11436/api/tags" -UseBasicParsing | Out-Null
            $ready = $true
            break
        } catch {
            Start-Sleep -Seconds 1
        }
    }
    if (-not $ready) {
        throw "The temporary Ollama server did not become ready."
    }
    & $ollamaPath pull $Model
    if ($LASTEXITCODE -ne 0) {
        throw "Ollama could not download $Model."
    }
} finally {
    if ($buildServer -and -not $buildServer.HasExited) {
        & taskkill.exe /PID $buildServer.Id /T /F > $null 2>&1
    }
}

if (-not (Test-Path (Join-Path $modelsSource "blobs"))) {
    throw "Ollama model files were not found in $modelsSource."
}

$bundle = Join-Path $projectRoot "dist\Inv_Reader"
if (Test-Path $bundle) {
    Remove-Item $bundle -Recurse -Force
}

Write-Host "Building Inv_Reader.exe..."
$activationModule = Join-Path $projectRoot "src\activation_build_config.py"
$activationServerLiteral = $ActivationServer.TrimEnd('/').Replace("'", "''")
$activationPublicKeyLiteral = $ActivationPublicKey.Trim().Replace("'", "''")
@"
ACTIVATION_SERVER = '$activationServerLiteral'
ACTIVATION_PUBLIC_KEY = '$activationPublicKeyLiteral'
"@ | Set-Content $activationModule -Encoding UTF8
try {
    & $Python -m PyInstaller --noconfirm --clean --onedir --windowed `
        --name Inv_Reader `
        --hidden-import fitz `
        --hidden-import pypdf `
        --hidden-import docx `
        --hidden-import openpyxl `
        --hidden-import src.activation_build_config `
        src\summarize_documents.py
    if ($LASTEXITCODE -ne 0) {
        throw "PyInstaller failed."
    }
} finally {
    Remove-Item $activationModule -Force -ErrorAction SilentlyContinue
}

$ollamaDirectory = Join-Path $bundle "ollama"
$bundleModels = Join-Path $bundle "models"
New-Item $ollamaDirectory -ItemType Directory -Force | Out-Null
New-Item $bundleModels -ItemType Directory -Force | Out-Null
Copy-Item (Join-Path $ollamaInstall "*") $ollamaDirectory -Recurse -Force
Copy-Item (Join-Path $modelsSource "*") $bundleModels -Recurse -Force
Copy-Item (Join-Path $projectRoot "Keys.pdf") $bundle -Force
@{
    model = $Model
    activation_required = $true
    activation_server = $ActivationServer.TrimEnd('/')
    activation_public_key = $ActivationPublicKey.Trim()
    product = "Inv_Reader"
} | ConvertTo-Json | Set-Content (Join-Path $bundle "portable_config.json") -Encoding UTF8
New-Item (Join-Path $bundle "input") -ItemType Directory -Force | Out-Null
New-Item (Join-Path $bundle "output") -ItemType Directory -Force | Out-Null

Write-Host "Portable application created at $bundle"
Write-Host "Copy the entire Inv_Reader directory to a Windows machine, then run Inv_Reader.exe."
