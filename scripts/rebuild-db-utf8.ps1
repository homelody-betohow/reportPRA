# 清空数据卷并按 UTF-8 重新初始化（解决备注乱码根因）
$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $PSScriptRoot
Set-Location $Root

Write-Host "将删除 MySQL 数据卷并重建库（所有表数据会清空）"
$confirm = Read-Host "输入 yes 继续"
if ($confirm -ne "yes") { exit 0 }

docker compose down -v
docker compose up -d
Write-Host "等待初始化..."
Start-Sleep -Seconds 35
python "$Root\scripts\fix_table_comments.py"
Write-Host "完成。Navicat 请断开重连后查看。"
