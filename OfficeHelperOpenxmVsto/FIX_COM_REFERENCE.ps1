# COM 引用修复脚本
# 此脚本帮助修复 Microsoft.Office.Interop.PowerPoint 的 COM 引用问题

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "COM 引用修复脚本" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# 检查 Visual Studio 安装
Write-Host "检查 Visual Studio 安装..." -ForegroundColor Yellow

$vsPaths = @(
    "C:\Program Files\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe",
    "C:\Program Files\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe",
    "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe",
    "C:\Program Files (x86)\Microsoft Visual Studio\2019\Enterprise\MSBuild\Current\Bin\MSBuild.exe",
    "C:\Program Files (x86)\Microsoft Visual Studio\2019\Professional\MSBuild\Current\Bin\MSBuild.exe",
    "C:\Program Files (x86)\Microsoft Visual Studio\2019\Community\MSBuild\Current\Bin\MSBuild.exe",
    "C:\Program Files (x86)\Microsoft Visual Studio\2017\Enterprise\MSBuild\15.0\Bin\MSBuild.exe",
    "C:\Program Files (x86)\Microsoft Visual Studio\2017\Professional\MSBuild\15.0\Bin\MSBuild.exe",
    "C:\Program Files (x86)\Microsoft Visual Studio\2017\Community\MSBuild\15.0\Bin\MSBuild.exe"
)

$msbuildPath = $null
foreach ($path in $vsPaths) {
    if (Test-Path $path) {
        $msbuildPath = $path
        Write-Host "找到 MSBuild: $path" -ForegroundColor Green
        break
    }
}

if (-not $msbuildPath) {
    Write-Host "错误: 未找到 Visual Studio MSBuild" -ForegroundColor Red
    Write-Host ""
    Write-Host "解决方案:" -ForegroundColor Yellow
    Write-Host "1. 安装 Visual Studio 2019 或 2022" -ForegroundColor White
    Write-Host "2. 确保安装了 '.NET Framework 4.8 开发工具' 工作负载" -ForegroundColor White
    Write-Host "3. 或者使用 Visual Studio IDE 打开项目并构建" -ForegroundColor White
    exit 1
}

# 检查 Office 安装
Write-Host ""
Write-Host "检查 Microsoft Office 安装..." -ForegroundColor Yellow

$officePaths = @(
    "C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE",
    "C:\Program Files (x86)\Microsoft Office\root\Office16\POWERPNT.EXE",
    "C:\Program Files\Microsoft Office\Office16\POWERPNT.EXE",
    "C:\Program Files (x86)\Microsoft Office\Office16\POWERPNT.EXE"
)

$officeInstalled = $false
foreach ($path in $officePaths) {
    if (Test-Path $path) {
        Write-Host "找到 PowerPoint: $path" -ForegroundColor Green
        $officeInstalled = $true
        break
    }
}

if (-not $officeInstalled) {
    Write-Host "警告: 未找到 Microsoft PowerPoint" -ForegroundColor Yellow
    Write-Host "请确保已安装 Microsoft Office 2016 或更高版本" -ForegroundColor Yellow
}

# 获取项目路径
$projectPath = Join-Path $PSScriptRoot "OfficeHelperOpenXml.csproj"
if (-not (Test-Path $projectPath)) {
    Write-Host "错误: 未找到项目文件: $projectPath" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "项目路径: $projectPath" -ForegroundColor Cyan

# 清理项目
Write-Host ""
Write-Host "清理项目..." -ForegroundColor Yellow
& $msbuildPath $projectPath /t:Clean /p:Configuration=Debug /verbosity:minimal
if ($LASTEXITCODE -ne 0) {
    Write-Host "清理失败，但继续..." -ForegroundColor Yellow
}

# 恢复 NuGet 包
Write-Host ""
Write-Host "恢复 NuGet 包..." -ForegroundColor Yellow
& $msbuildPath $projectPath /t:Restore /verbosity:minimal
if ($LASTEXITCODE -ne 0) {
    Write-Host "NuGet 恢复失败，但继续..." -ForegroundColor Yellow
}

# 构建项目
Write-Host ""
Write-Host "构建项目（使用 Visual Studio MSBuild）..." -ForegroundColor Yellow
Write-Host "注意: 如果 COM 引用仍然失败，请在 Visual Studio 中手动添加引用" -ForegroundColor Yellow
Write-Host ""

& $msbuildPath $projectPath /t:Build /p:Configuration=Debug /verbosity:normal

if ($LASTEXITCODE -eq 0) {
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Green
    Write-Host "构建成功！" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
} else {
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Red
    Write-Host "构建失败" -ForegroundColor Red
    Write-Host "========================================" -ForegroundColor Red
    Write-Host ""
    Write-Host "如果仍然出现 COM 引用错误，请执行以下步骤:" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "1. 在 Visual Studio 中打开项目" -ForegroundColor White
    Write-Host "2. 右键点击项目 → 添加 → 引用" -ForegroundColor White
    Write-Host "3. 选择 COM 选项卡" -ForegroundColor White
    Write-Host "4. 勾选以下项:" -ForegroundColor White
    Write-Host "   - Microsoft Office 16.0 Object Library" -ForegroundColor White
    Write-Host "   - Microsoft PowerPoint 16.0 Object Library" -ForegroundColor White
    Write-Host "5. 点击确定" -ForegroundColor White
    Write-Host "6. 清理并重新构建解决方案" -ForegroundColor White
    Write-Host ""
    Write-Host "详细说明请参考: COM_REFERENCE_FIX.md" -ForegroundColor Cyan
    exit 1
}

