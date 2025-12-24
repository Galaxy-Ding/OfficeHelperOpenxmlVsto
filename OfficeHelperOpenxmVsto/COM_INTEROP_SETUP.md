# COM 互操作程序集设置指南

## 概述

本项目已升级到 .NET 8.0，并使用 NuGet 包引用 Office 互操作程序集。但是，`Microsoft.Office.Core` 需要通过 `tlbimp` 工具从 COM 类型库生成。

## 设置步骤

### 方法 1：使用 tlbimp 工具生成互操作程序集（推荐）

1. **找到 MSO.DLL 文件位置**
   - 通常位于：`C:\Program Files\Common Files\Microsoft Shared\OFFICE16\MSO.DLL`
   - 或者：`C:\Program Files (x86)\Common Files\Microsoft Shared\OFFICE16\MSO.DLL`

2. **找到 tlbimp 工具**
   - 通常位于：`C:\Program Files (x86)\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8 Tools\tlbimp.exe`
   - 或者使用 Visual Studio Developer Command Prompt 中的 tlbimp

3. **生成互操作程序集**
   ```powershell
   # 创建 Interop 目录
   mkdir OfficeHelperOpenxmVsto\Interop
   
   # 生成互操作程序集
   tlbimp "C:\Program Files\Common Files\Microsoft Shared\OFFICE16\MSO.DLL" /out:"OfficeHelperOpenxmVsto\Interop\Microsoft.Office.Core.dll" /namespace:Microsoft.Office.Core
   ```

4. **验证生成的文件**
   - 确认 `OfficeHelperOpenxmVsto\Interop\Microsoft.Office.Core.dll` 文件已创建

### 方法 2：使用 Visual Studio

1. 在 Visual Studio 中打开项目
2. 右键点击项目 → 添加 → 引用
3. 选择 COM 选项卡
4. 找到并勾选：
   - `Microsoft Office 16.0 Object Library` (Microsoft.Office.Core)
   - `Microsoft PowerPoint 16.0 Object Library` (Microsoft.Office.Interop.PowerPoint)
5. 点击确定

注意：Visual Studio 会自动生成互操作程序集到 `obj` 目录。

### 方法 3：使用已安装的 PIA（Primary Interop Assemblies）

如果 Office 已安装 PIA，可以直接引用：

1. 查找 PIA 位置（通常在 GAC 或 Office 安装目录）
2. 在项目文件中添加引用

## 当前配置

项目已配置为：
- 使用 NuGet 包 `Microsoft.Office.Interop.PowerPoint` (15.0.4420.1018)
- 自动查找 `Interop\Microsoft.Office.Core.dll`
- 如果找不到，构建时会显示警告信息

## 验证设置

运行以下命令验证设置是否正确：

```powershell
dotnet build OfficeHelperOpenxmVsto\OfficeHelperOpenXml.csproj
```

如果看到关于 `Microsoft.Office.Core` 的编译错误，请按照上述步骤生成互操作程序集。

## 故障排除

### 问题：找不到 tlbimp 工具

**解决方案**：
1. 安装 Visual Studio 或 Windows SDK
2. 使用 Visual Studio Developer Command Prompt
3. 或者从 [Microsoft 下载中心](https://www.microsoft.com/download) 下载 Windows SDK

### 问题：tlbimp 生成失败

**可能原因**：
- Office 未正确安装
- MSO.DLL 文件损坏
- 权限不足

**解决方案**：
- 确保 Office 已正确安装
- 以管理员身份运行命令
- 检查文件路径是否正确

### 问题：编译时仍然找不到 Microsoft.Office.Core

**解决方案**：
1. 确认 `Interop\Microsoft.Office.Core.dll` 文件存在
2. 检查项目文件中的路径配置
3. 清理并重新构建项目：`dotnet clean && dotnet build`

## 参考资源

- [.NET 中的 COM 互操作](https://docs.microsoft.com/dotnet/standard/native-interop/cominterop)
- [tlbimp 工具文档](https://docs.microsoft.com/dotnet/framework/tools/tlbimp-exe-type-library-importer)
- [Office 主互操作程序集](https://docs.microsoft.com/visualstudio/vsto/office-primary-interop-assemblies)

