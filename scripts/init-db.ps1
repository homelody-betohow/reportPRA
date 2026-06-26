# 在已运行的容器上重新执行 schema（仅当需要手动修复库表时使用）
$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $PSScriptRoot
Set-Location $Root

$envFile = Join-Path $Root ".env"
$rootPass = "report_root"
if (Test-Path $envFile) {
    Get-Content $envFile | ForEach-Object {
        if ($_ -match '^MYSQL_ROOT_PASSWORD=(.+)$') { $rootPass = $matches[1].Trim() }
    }
}

$sql = Join-Path $Root "database\schema.sql"
Write-Host "导入 $sql ..."
Get-Content $sql -Raw -Encoding UTF8 | docker exec -i report-mysql mysql -uroot "-p$rootPass"
Write-Host "完成。"
