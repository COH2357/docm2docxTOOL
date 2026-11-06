# Build the portable runtime (runtime\python + runtime\everything)
#
# Copies this machine's Python into a trimmed, self-contained portable copy
# (with tkinter and pywin32), and copies the Everything binary into
# runtime\everything, so the whole folder runs on any Windows machine
# without installing Python / pywin32 / Everything.
#
# NOTE: keep this file ASCII-only. Windows PowerShell 5.1 reads .ps1 files
# as ANSI unless they carry a UTF-8 BOM. This file must be saved WITH BOM.
$ErrorActionPreference = 'Stop'
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$Runtime = Join-Path $Root 'runtime'
$PyTarget = Join-Path $Runtime 'python'
$EvTarget = Join-Path $Runtime 'everything'

function Write-Step($text) { Write-Host "==> $text" -ForegroundColor Cyan }
function Write-Info($text) { Write-Host "    $text" }

# --------------------------------------------------------------------------- #
# 1. locate the host Python
# --------------------------------------------------------------------------- #
Write-Step "Find host Python"
$source = $null
$launcher = Get-Command py.exe -ErrorAction SilentlyContinue
if ($launcher) {
    try {
        $path = & $launcher.Source -3 -c "import sys; print(sys.executable)" 2>$null
        if ($path -and (Test-Path $path)) { $source = Split-Path -Parent $path }
    } catch { }
}
if (-not $source) {
    $python = Get-Command python.exe -ErrorAction SilentlyContinue
    if ($python) { $source = Split-Path -Parent $python.Source }
}
if (-not $source -or -not (Test-Path (Join-Path $source 'python.exe'))) {
    throw "Python not found. Install Python 3.8+ first (tick 'Add Python to PATH')."
}
Write-Info "Source: $source"

if (Test-Path $PyTarget) { Remove-Item -Recurse -Force $PyTarget }
New-Item -ItemType Directory -Force -Path $PyTarget | Out-Null

# --------------------------------------------------------------------------- #
# 2. copy the interpreter
# --------------------------------------------------------------------------- #
Write-Step "Copy Python runtime"
Get-ChildItem $source -File | Where-Object {
    $_.Name -match '^(python.*\.(exe|dll)|vcruntime.*\.dll|.*\.pyd)$'
} | ForEach-Object {
    Copy-Item $_.FullName (Join-Path $PyTarget $_.Name) -Force
}
foreach ($dir in @('DLLs', 'tcl')) {
    if (Test-Path (Join-Path $source $dir)) {
        Copy-Item (Join-Path $source $dir) (Join-Path $PyTarget $dir) -Recurse -Force
        Write-Info "copied $dir"
    }
}

# stdlib, without the test suite / IDLE / third-party packages
Write-Step "Copy standard library"
$libSrc = Join-Path $source 'Lib'
$libDst = Join-Path $PyTarget 'Lib'
New-Item -ItemType Directory -Force -Path $libDst | Out-Null
$skipDirs = @('site-packages', 'test', 'idlelib', 'turtledemo', '__pycache__', 'lib2to3')
Get-ChildItem $libSrc -Directory | Where-Object { $skipDirs -notcontains $_.Name } | ForEach-Object {
    Copy-Item $_.FullName (Join-Path $libDst $_.Name) -Recurse -Force
}
Get-ChildItem $libSrc -File | ForEach-Object {
    Copy-Item $_.FullName (Join-Path $libDst $_.Name) -Force
}
Write-Info "standard library copied"

# --------------------------------------------------------------------------- #
# 3. site-packages: pywin32 + pip only
# --------------------------------------------------------------------------- #
Write-Step "Copy pywin32 and pip"
$spSrc = Join-Path $libSrc 'site-packages'
$spDst = Join-Path $libDst 'site-packages'
New-Item -ItemType Directory -Force -Path $spDst | Out-Null
$keepDirs = @(
    'win32', 'win32com', 'win32comext', 'pythonwin', 'pywin32_system32',
    'win32ctypes', 'pywin32_ctypes', 'pip', 'setuptools', '_distutils_hack'
)
foreach ($name in $keepDirs) {
    $p = Join-Path $spSrc $name
    if (Test-Path $p) {
        Copy-Item $p (Join-Path $spDst $name) -Recurse -Force
    }
}
Get-ChildItem $spSrc -File | Where-Object {
    $_.Name -match '^(pythoncom\.py|pywin32.*|distutils-precedence\.pth)$'
} | ForEach-Object { Copy-Item $_.FullName (Join-Path $spDst $_.Name) -Force }
Write-Info "pywin32 + pip copied"

Get-ChildItem $PyTarget -Recurse -Directory -Filter '__pycache__' -ErrorAction SilentlyContinue |
    Remove-Item -Recurse -Force -ErrorAction SilentlyContinue

# --------------------------------------------------------------------------- #
# 4. Everything
# --------------------------------------------------------------------------- #
Write-Step "Collect Everything"
if (Test-Path $EvTarget) { Remove-Item -Recurse -Force $EvTarget }
New-Item -ItemType Directory -Force -Path $EvTarget | Out-Null

$evSource = $null
foreach ($cand in @('C:\Program Files\Everything', 'C:\Program Files (x86)\Everything')) {
    if (Test-Path (Join-Path $cand 'Everything.exe')) { $evSource = $cand; break }
}
if ($evSource) {
    foreach ($f in @('Everything.exe', 'Everything.lng', 'License.txt')) {
        $p = Join-Path $evSource $f
        if (Test-Path $p) { Copy-Item $p (Join-Path $EvTarget $f) -Force }
    }
    Write-Info "copied Everything.exe from $evSource"
} else {
    Write-Info "Everything is not installed locally, downloading from voidtools.com ..."
    try {
        $page = Invoke-WebRequest -Uri 'https://www.voidtools.com/downloads/' -UseBasicParsing -TimeoutSec 30
        $m = [regex]::Match($page.Content, 'Everything-[\d\.]+\.zip')
        if ($m.Success) {
            $url = "https://www.voidtools.com/$($m.Value)"
            $zip = Join-Path $env:TEMP $m.Value
            Invoke-WebRequest -Uri $url -OutFile $zip -UseBasicParsing -TimeoutSec 180
            Expand-Archive -Path $zip -DestinationPath $EvTarget -Force
            Get-ChildItem $EvTarget -Recurse -Filter 'Everything*.exe' | ForEach-Object {
                Copy-Item $_.FullName (Join-Path $EvTarget 'Everything.exe') -Force
            }
            Remove-Item $zip -Force
            Write-Info "downloaded $($m.Value)"
        } else {
            Write-Info "no download link found, skipping (falls back to Windows index)"
        }
    } catch {
        Write-Info "download failed: $($_.Exception.Message)"
    }
}

# --------------------------------------------------------------------------- #
# 5. self test
# --------------------------------------------------------------------------- #
Write-Step "Verify portable runtime"
$pyExe = Join-Path $PyTarget 'python.exe'
if (-not (Test-Path $pyExe)) { throw "portable python was not created" }
$check = & $pyExe -c "import sys,tkinter,win32com.client,pythoncom;print('OK',sys.version.split()[0],tkinter.TkVersion)" 2>&1
Write-Info "self test: $check"
if ($check -notmatch '^OK') { throw "portable python self test failed" }

$size = (Get-ChildItem $PyTarget -Recurse -File | Measure-Object -Sum Length).Sum
Write-Host ""
Write-Host "Portable runtime ready:" -ForegroundColor Green
Write-Host ("    runtime\python      {0:N0} MB" -f ($size / 1MB))
Write-Host ("    runtime\everything  {0:N0} MB" -f ((Get-ChildItem $EvTarget -Recurse -File | Measure-Object -Sum Length).Sum / 1MB))
Write-Host "    run.bat now uses the bundled runtime automatically."
