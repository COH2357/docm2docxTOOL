#!/usr/bin/env pwsh
# 创建 GitHub Release 并上传打包好的 zip 附件。
#
# 用法:
#   pwsh -File tools/release.ps1 -Version 1.0.0 -NotesFile tools/release-notes-v1.0.0.md
#
# 令牌获取顺序(任一即可):
#   1. -Token 参数
#   2. 环境变量 GITHUB_TOKEN
#   3. gh CLI: gh auth token
#   4. 剪贴板(手动复制 40 位 PAT 后运行)
#
# 需要仓库 admin 权限。脚本幂等:同名 Release 复用,同名附件跳过。
param(
  [string]$Version = '1.0.0',
  [string]$Repo = 'COH2357/docm2docxTOOL',
  [string]$NotesFile = '',
  [string]$Token = ''
)

$ErrorActionPreference = 'Stop'

if (-not $Token) { $Token = $env:GITHUB_TOKEN }
if (-not $Token) {
  $gh = Get-Command gh -ErrorAction SilentlyContinue
  if ($gh) { $Token = (& gh auth token 2>$null) }
}
if (-not $Token) {
  $clip = Get-Clipboard -Raw -ErrorAction SilentlyContinue
  if ($clip) { $Token = $clip.Trim() }
}
if (-not $Token) { throw '没有可用令牌:请用 -Token / GITHUB_TOKEN / gh auth token / 剪贴板' }

$h = @{
  Authorization            = "Bearer $Token"
  Accept                   = 'application/vnd.github+json'
  'X-GitHub-Api-Version'   = '2022-11-28'
  'User-Agent'             = 'docm2docx-release'
}

$tag = "v$Version"

# 发布说明:未指定文件时用仓库根目录 README 的摘要
if ($NotesFile -and (Test-Path $NotesFile)) {
  $notes = [System.IO.File]::ReadAllText((Resolve-Path $NotesFile), [System.Text.Encoding]::UTF8)
} else {
  $notes = "DOCM → DOCX 批量无损转换器 $tag"
}

$payload = @{
  tag_name        = $tag
  name            = "$tag · DOCM 批量无损转换为 DOCX"
  body            = $notes
  draft           = $false
  prerelease      = $false
  target_commitish = 'main'
} | ConvertTo-Json -Depth 5

Write-Host "==> 创建/获取 Release $tag …"
$release = $null
try {
  $release = Invoke-RestMethod -Uri "https://api.github.com/repos/$Repo/releases/tags/$tag" `
    -Headers $h -TimeoutSec 60
  Write-Host "    复用已有: $($release.html_url)"
} catch {
  $release = Invoke-RestMethod -Uri "https://api.github.com/repos/$Repo/releases" `
    -Method Post -Headers $h `
    -Body ([System.Text.Encoding]::UTF8.GetBytes($payload)) `
    -ContentType 'application/json; charset=utf-8' -TimeoutSec 60
  Write-Host "    已创建: $($release.html_url)"
}

$uploadUrl = $release.upload_url -replace '\{.*\}$',''
$assets = @(
  (Join-Path (Split-Path -Parent $PSScriptRoot) "docm2docx-v$Version-portable.zip"),
  (Join-Path (Split-Path -Parent $PSScriptRoot) "docm2docx-v$Version-source.zip")
)

foreach ($path in $assets) {
  if (-not (Test-Path $path)) { Write-Host "==> 跳过(不存在): $path"; continue }
  $name = [System.IO.Path]::GetFileName($path)
  $existingAsset = @($release.assets) | Where-Object { $_.name -eq $name } | Select-Object -First 1
  if ($existingAsset) { Write-Host "==> 跳过(已存在): $name"; continue }
  $size = [int]((Get-Item $path).Length / 1MB)
  Write-Host "==> 上传 $name ($size MB) …"
  # 注意:必须用 ${uploadUrl} 界定变量名,否则 PowerShell 会把 ? 当成变量名的一部分
  $url = "${uploadUrl}?name=$([uri]::EscapeDataString($name))"
  $asset = Invoke-RestMethod -Uri $url -Method Post -Headers $h `
    -ContentType 'application/octet-stream' -InFile $path -TimeoutSec 900
  Write-Host "    完成: $($asset.browser_download_url)"
}

Write-Host ''
Write-Host '==> 校验 …'
$final = Invoke-RestMethod -Uri "https://api.github.com/repos/$Repo/releases/tags/$tag" `
  -Headers $h -TimeoutSec 60
Write-Host "    标题    : $($final.name)"
Write-Host "    标签    : $($final.tag_name)"
Write-Host "    草稿    : $($final.draft)   预发布: $($final.prerelease)"
Write-Host "    附件    : $(@($final.assets).Count) 个"
foreach ($a in $final.assets) {
  Write-Host ("      - {0}  {1:N1} MB" -f $a.name, ($a.size / 1MB))
}
Write-Host "    地址    : $($final.html_url)"
