# 在已运行的 MySQL 容器上导入 docs/database 下的 DDL（按需使用）
$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $PSScriptRoot
Set-Location $Root

$envFile = Join-Path $Root ".env"
$containerName = "rpa-task-mysql"
$rootPass = "root_secret"
if (Test-Path $envFile) {
    Get-Content $envFile | ForEach-Object {
        if ($_ -match '^MYSQL_CONTAINER_NAME=(.+)$') { $containerName = $matches[1].Trim() }
        if ($_ -match '^MYSQL_ROOT_PASSWORD=(.+)$') { $rootPass = $matches[1].Trim() }
    }
}

$sqlDir = Join-Path $Root "docs\database"
if (-not (Test-Path $sqlDir)) {
    Write-Error "未找到 $sqlDir"
}

$files = Get-ChildItem $sqlDir -Filter "*.sql" | Sort-Object Name
if (-not $files) {
    Write-Error "docs\database 下没有 .sql 文件"
}

Write-Host "将向容器 $containerName 导入 $($files.Count) 个 SQL 文件..."
foreach ($f in $files) {
    Write-Host "  -> $($f.Name)"
    Get-Content $f.FullName -Raw -Encoding UTF8 | docker exec -i $containerName mysql -uroot "-p$rootPass"
}
Write-Host "完成。"
