# 下单费用核算工具 打包脚本
# 修改 code\order_cost.py 后，关闭正在运行的程序，然后在 PowerShell 里执行本脚本即可重新生成 exe。
# 用法：  powershell -ExecutionPolicy Bypass -File "D:\Workspace\xc\order_cost_accounting\exe\build.ps1"
# 注意：本文件必须保存为 UTF-8 带 BOM，否则 Windows PowerShell 5.1 会把中文名解析成乱码。

$ErrorActionPreference = 'Stop'

$Root   = 'D:\Workspace\xc\order_cost_accounting'
$EnvDir = 'D:\miniconda3\envs\order_cost'
$Source = Join-Path $Root 'code\order_cost.py'
$Dist   = Join-Path $Root 'exe'
$Build  = Join-Path $Root 'build'
$Name   = '下单费用核算工具'

foreach ($p in @($EnvDir, $Source, $Dist)) {
    if (-not (Test-Path $p)) {
        Write-Host "[错误] 路径不存在：$p" -ForegroundColor Red
        exit 1
    }
}

# 程序未关闭时 PyInstaller 无法覆盖 exe，提前拦下而不是等打包到一半失败
if (Get-Process -Name $Name -ErrorAction SilentlyContinue) {
    Write-Host "[提示] 检测到 $Name.exe 正在运行，请先关闭后再打包。" -ForegroundColor Yellow
    exit 1
}

# conda 环境未激活时 Library\bin 不在 PATH，PyInstaller 会漏收 tcl86t/tk86t/libexpat 等 DLL，
# 打包后运行会报 _tkinter / pyexpat 的 DLL load failed。脚本结束前恢复原 PATH，避免污染当前会话。
$oldPath = $env:Path
$env:Path = "$EnvDir;$EnvDir\Library\bin;$EnvDir\Scripts;$EnvDir\Library\usr\bin;$oldPath"

try {
    Write-Host "开始打包：$Source -> $Dist\$Name.exe" -ForegroundColor Cyan
    # --clean 必须保留：否则 PyInstaller 复用上次的依赖分析缓存，漏收的 DLL 不会补上
    $ErrorActionPreference = 'Continue'
    & "$EnvDir\python.exe" -m PyInstaller --noconfirm --clean --onefile --windowed `
        --name $Name `
        --distpath $Dist `
        --workpath $Build `
        --specpath $Build `
        $Source
    $code = $LASTEXITCODE
    $ErrorActionPreference = 'Stop'

    if ($code -ne 0) {
        Write-Host "[错误] PyInstaller 退出码 $code，打包失败" -ForegroundColor Red
        exit $code
    }

    $exe = Join-Path $Dist "$Name.exe"
    if (-not (Test-Path $exe)) {
        Write-Host "[错误] 打包结束但未找到 $exe" -ForegroundColor Red
        exit 1
    }
    $f = Get-Item $exe
    Write-Host ("[完成] {0}`n       大小 {1:N2} MB   时间 {2}" -f $f.FullName, ($f.Length / 1MB), $f.LastWriteTime) -ForegroundColor Green
}
finally {
    $env:Path = $oldPath
}
