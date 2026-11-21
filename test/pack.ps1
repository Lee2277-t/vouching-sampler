# 抽凭助手 · 打发布包
# 用法： powershell -NoProfile -ExecutionPolicy Bypass -File test\pack.ps1
# 产出： ..\抽凭助手-v<版本>.zip，只含运行必需的文件（不含测试、不含本机路径）
$ErrorActionPreference = "Stop"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$root = Split-Path -Parent $here
$outDir = Split-Path -Parent $root

# 版本从 core.js 里读，避免手写两处
$core = Get-Content -Raw -Encoding UTF8 (Join-Path $root "core.js")
if ($core -notmatch "var VERSION = '([0-9.]+)'") { throw "core.js 里找不到 VERSION" }
$version = $Matches[1]
$zip = Join-Path $outDir ("抽凭助手-v{0}.zip" -f $version)
$stage = Join-Path $env:TEMP ("chouping-pack-" + [guid]::NewGuid().ToString("N").Substring(0, 8))
$inner = Join-Path $stage ("抽凭助手-v{0}" -f $version)
New-Item -ItemType Directory -Path $inner -Force | Out-Null

# 运行必需的：界面三件 + 核心 + 本地库 + 文档 + 许可
$files = @("index.html", "app.js", "app.css", "core.js", "README.md", "LICENSE")
foreach ($f in $files) {
    $src = Join-Path $root $f
    if (-not (Test-Path $src)) { throw "缺文件：$f" }
    Copy-Item $src (Join-Path $inner $f)
}
Copy-Item (Join-Path $root "vendor") (Join-Path $inner "vendor") -Recurse

# 样例目录（有就带上，方便对照）
$samples = Join-Path $root "samples"
if (Test-Path $samples) { Copy-Item $samples (Join-Path $inner "samples") -Recurse }

if (Test-Path $zip) { Remove-Item $zip -Force }
Compress-Archive -Path $inner -DestinationPath $zip -CompressionLevel Optimal

$size = [math]::Round((Get-Item $zip).Length / 1KB, 1)
$count = (Get-ChildItem $inner -Recurse -File).Count
Write-Host ("已打包：{0}" -f $zip)
Write-Host ("  版本 {0} · {1} 个文件 · {2} KB" -f $version, $count, $size)
Write-Host ("  暂存目录（可对账用）：{0}" -f $inner)
