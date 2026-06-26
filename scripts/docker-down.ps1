$Root = Split-Path -Parent $PSScriptRoot
Set-Location $Root
docker compose down
Write-Host "已停止 MySQL 容器（数据卷 report_mysql_data 仍保留）"
