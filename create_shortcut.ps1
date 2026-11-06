# Create a Desktop shortcut for the DOCM -> DOCX converter.
# ASCII-only on purpose; the Chinese shortcut name comes from shortcut_name.txt.
$ErrorActionPreference = 'Stop'

$appDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$target = Join-Path $appDir 'run.bat'
$nameFile = Join-Path $appDir 'shortcut_name.txt'

if (-not (Test-Path -LiteralPath $nameFile)) {
    throw "Missing file: $nameFile"
}

$name = (Get-Content -LiteralPath $nameFile -Encoding UTF8 -Raw).Trim()
if ([string]::IsNullOrWhiteSpace($name)) {
    $name = 'DOCM to DOCX Converter'
}

$desktop = [Environment]::GetFolderPath('Desktop')
if ([string]::IsNullOrWhiteSpace($desktop)) {
    throw 'Desktop folder not found'
}
New-Item -ItemType Directory -Path $desktop -Force | Out-Null

$shortcutPath = Join-Path $desktop ($name + '.lnk')
$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = $target
$shortcut.WorkingDirectory = $appDir
$shortcut.Description = 'DOCM batch converter (lossless to DOCX)'

$pythonw = Get-Command pythonw.exe -ErrorAction SilentlyContinue
if ($pythonw) {
    $shortcut.IconLocation = ($pythonw.Source + ',0')
}

$shortcut.Save()
Write-Host ("  [OK] Desktop shortcut created: " + $shortcutPath)
