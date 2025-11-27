# ----- Fix-Kernel.ps1 (ASCII-safe) -----
$ErrorActionPreference = "Stop"

Write-Host "Killing processes that may lock site-packages..."
$procs = "code","python","jupyter-lab","jupyter-notebook","node"
foreach ($p in $procs) {
  Start-Process cmd "/c taskkill /F /IM $p.exe" -WindowStyle Hidden -ErrorAction SilentlyContinue
}

# Stop OneDrive if running (prevents file locks)
$oneDriveWasRunning = $false
try {
  $od = Get-Process OneDrive -ErrorAction SilentlyContinue
  if ($od) {
    $oneDriveWasRunning = $true
    & "$env:LOCALAPPDATA\Microsoft\OneDrive\OneDrive.exe" /shutdown | Out-Null
  }
} catch {}

# Paths
$projectRoot = (Get-Location).Path
$venv       = Join-Path $projectRoot ".venv"
$venvPy     = Join-Path $venv "Scripts\python.exe"
$venvUv     = Join-Path $venv "Scripts\uv.exe"
$sitePkgs   = Join-Path $venv "Lib\site-packages"

if (!(Test-Path $venvPy)) { throw ".venv not found at $venv. Create it first (uv venv .venv)." }
if (!(Test-Path $sitePkgs)) { throw "site-packages not found at $sitePkgs" }

Write-Host "Clearing attributes and removing problematic dist-info folders..."
$targets = @(
  "langchain_community-*.dist-info",
  "langchain_openai-*.dist-info",
  "ipykernel-*.dist-info",
  "jupyter_client-*.dist-info",
  "jupyter_core-*.dist-info"
)

foreach ($pattern in $targets) {
  Get-ChildItem -Path $sitePkgs -Filter $pattern -Directory -ErrorAction SilentlyContinue | ForEach-Object {
    try {
      attrib -R -S -H $_.FullName /S | Out-Null
      Remove-Item -Recurse -Force $_.FullName -ErrorAction SilentlyContinue
      Write-Host ("  removed " + $_.Name)
    } catch {
      Write-Host ("  could not remove " + $_.FullName + " : " + $_.Exception.Message)
    }
  }
}

Write-Host "Reinstalling Jupyter kernel components..."
$env:UV_LINK_MODE = "copy"   # safer on OneDrive

# Choose uv inside venv if present; else fall back to global
$uvCmd = "uv"
if (Test-Path $venvUv) { $uvCmd = $venvUv }

# Install/upgrade kernel deps
& $uvCmd add -U jupyter ipykernel jupyter_core jupyter_client

Write-Host "Registering kernel 'RAG Kernel'..."
& $venvPy -m ipykernel install --user --name RAG-venv --display-name "RAG Kernel"

# Restart OneDrive if it was running
if ($oneDriveWasRunning) {
  try { Start-Process "$env:LOCALAPPDATA\Microsoft\OneDrive\OneDrive.exe" | Out-Null } catch {}
}

Write-Host ""
Write-Host "Done. Reopen VS Code and select the kernel 'RAG Kernel'."
