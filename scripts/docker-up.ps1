# 启动 Docker MySQL 并等待就绪
$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $PSScriptRoot
Set-Location $Root

if (-not (Test-Path ".env")) {
    Copy-Item ".env.example" ".env"
    Write-Host "已根据 .env.example 创建 .env"
}

Write-Host "正在启动 MySQL 容器..."
docker compose up -d

Write-Host "等待 MySQL 健康检查（最多约 60 秒）..."
$max = 30
for ($i = 1; $i -le $max; $i++) {
    $status = docker inspect --format='{{.State.Health.Status}}' report-mysql 2>$null
    if ($status -eq "healthy") {
        Write-Host "MySQL 已就绪。"
        $db = "bth-report"
        $user = "betohow"
        if (Test-Path (Join-Path $Root ".env")) {
            Get-Content (Join-Path $Root ".env") | ForEach-Object {
                if ($_ -match '^MYSQL_DATABASE=(.+)$') { $db = $matches[1].Trim() }
                if ($_ -match '^MYSQL_USER=(.+)$') { $user = $matches[1].Trim() }
            }
        }
        Write-Host "连接: 127.0.0.1:3307  用户: $user  库: $db"
        exit 0
    }
    Start-Sleep -Seconds 2
}

Write-Host "容器已启动，健康检查未通过，请执行: docker compose logs mysql"
exit 1
