# 构建说明

## ⚠️ 重要提示

本项目使用 **.NET Framework 4.8** 和 **COM 引用**，因此**不能使用 `dotnet build` 命令**。

必须使用 **Visual Studio MSBuild** 或 **Visual Studio IDE** 来构建项目。

## 为什么不能使用 `dotnet build`？

`.NET Core` 版本的 MSBuild 不支持 `ResolveComReference`，这是处理 COM 引用所必需的。只有 **.NET Framework** 版本的 MSBuild 支持 COM 引用。

## 构建方法

### 方法 1：使用 Visual Studio IDE（推荐）

1. **打开 Visual Studio 2019 或 2022**
2. **打开解决方案**
   - 文件 → 打开 → 项目/解决方案
   - 选择 `OfficeHelperOpenxmVsto.sln` 或 `OfficeHelperOpenXml.csproj`
3. **添加 COM 引用**（如果尚未添加）
   - 右键点击 `OfficeHelperOpenXml` 项目 → 添加 → 引用
   - 选择 **COM** 选项卡
   - 勾选：
     - ✅ `Microsoft Office 16.0 Object Library` (Microsoft.Office.Core)
     - ✅ `Microsoft PowerPoint 16.0 Object Library` (Microsoft.Office.Interop.PowerPoint)
   - 点击 **确定**
4. **构建项目**
   - 生成 → 生成解决方案 (Ctrl+Shift+B)
   - 或：生成 → 重新生成解决方案

### 方法 2：使用 Visual Studio MSBuild（命令行）

#### Windows PowerShell

```powershell
# 运行修复脚本（自动查找 MSBuild）
.\FIX_COM_REFERENCE.ps1

# 或手动使用 MSBuild
# 找到 Visual Studio MSBuild 路径（通常在以下位置之一）:
# C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe
# C:\Program Files (x86)\Microsoft Visual Studio\2019\Community\MSBuild\Current\Bin\MSBuild.exe

$msbuild = "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe"
& $msbuild OfficeHelperOpenXml.csproj /t:Build /p:Configuration=Debug
```

#### 命令提示符 (CMD)

```cmd
REM 运行修复脚本
FIX_COM_REFERENCE.ps1

REM 或手动使用 MSBuild
"C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" OfficeHelperOpenXml.csproj /t:Build /p:Configuration=Debug
```

### 方法 3：使用 Developer Command Prompt

1. **打开 Visual Studio Developer Command Prompt**
   - 开始菜单 → Visual Studio 2019/2022 → Developer Command Prompt
2. **导航到项目目录**
   ```cmd
   cd D:\pythonf\c_sharp_project\OfficeHelperOpenxmVsto\OfficeHelperOpenxmVsto
   ```
3. **构建项目**
   ```cmd
   msbuild OfficeHelperOpenXml.csproj /t:Build /p:Configuration=Debug
   ```

## 常见错误和解决方案

### 错误 1: `.NET Core 版本的 MSBuild 不支持"ResolveComReference"`

**原因**: 使用了 `dotnet build` 命令

**解决方案**: 
- ✅ 使用 Visual Studio IDE 构建
- ✅ 使用 Visual Studio MSBuild（不是 dotnet CLI）
- ✅ 使用 Developer Command Prompt

### 错误 2: `无法获取类型库"91493440-5a91-11cf-8700-00aa0060263b"`

**原因**: COM 引用未正确配置

**解决方案**:
1. 在 Visual Studio 中手动添加 COM 引用（见方法 1）
2. 确保已安装 Microsoft PowerPoint
3. 参考 [COM_REFERENCE_FIX.md](COM_REFERENCE_FIX.md)

### 错误 3: `命名空间"Microsoft.Office"中不存在类型或命名空间名"Interop"`

**原因**: COM 互操作程序集未生成或未引用

**解决方案**:
1. 在 Visual Studio 中添加 COM 引用
2. 清理并重新构建项目
3. 检查 `obj` 目录中是否有生成的互操作程序集

## 验证构建

构建成功后，应该看到：

```
生成成功。
    0 个警告
    0 个错误
```

输出文件应该在：
- `bin\Debug\net48\OfficeHelperOpenXml.exe`
- `bin\Debug\net48\Microsoft.Office.Interop.PowerPoint.dll` (互操作程序集)

## 开发环境要求

- ✅ Visual Studio 2019 或 2022
- ✅ .NET Framework 4.8 SDK
- ✅ Microsoft Office 2016 或更高版本（包含 PowerPoint）
- ✅ Windows 操作系统

## 相关文档

- [COM 引用修复指南](COM_REFERENCE_FIX.md)
- [VSTO 快速开始](VSTO_QUICK_START.md)
- [架构确认文档](ARCHITECTURE_CONFIRMATION.md)

