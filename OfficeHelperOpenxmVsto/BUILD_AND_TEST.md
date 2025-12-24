# 构建和测试指南

## ⚠️ 重要提示

由于项目使用 **.NET Framework 4.8** 和 **COM 引用**，需要使用 **Visual Studio 的 MSBuild**（.NET Framework 版本）来构建项目。

**.NET Core 版本的 MSBuild 不支持 COM 引用**，因此不能使用 `dotnet build` 命令。

## 🚀 构建步骤

### 方法 1：使用 Visual Studio（推荐）

1. **打开解决方案**
   ```powershell
   # 在项目根目录
   start OfficeHelperOpenxmVsto.sln
   ```

2. **在 Visual Studio 中构建**
   - 菜单：**生成** → **生成解决方案** (Ctrl+Shift+B)
   - 或右键点击解决方案 → **生成解决方案**

3. **验证构建结果**
   - 检查 `OfficeHelperOpenxmVsto\bin\Release\net48\` 目录
   - 应该看到 `OfficeHelperOpenXml.exe` 和相关的 DLL 文件

### 方法 2：使用 MSBuild 命令行

如果已安装 Visual Studio，可以使用 MSBuild：

```powershell
# 查找 MSBuild 路径（根据你的 Visual Studio 版本调整）
$msbuild = "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe"

# 构建项目
& $msbuild OfficeHelperOpenxmVsto\OfficeHelperOpenXml.csproj /p:Configuration=Release /p:Platform="Any CPU"
```

## ✅ 架构确认

### 读取：使用 OpenXML SDK ✅

**实现位置**：`Core/Readers/PresentationReader.cs`

```csharp
using (var doc = PresentationDocument.Open(filePath, false))
{
    // 使用 OpenXML SDK 读取 PowerPoint 文件
    var presentationPart = doc.PresentationPart;
    // ...
}
```

**验证**：
- ✅ 使用 `DocumentFormat.OpenXml.Packaging.PresentationDocument`
- ✅ 无需 PowerPoint 应用程序运行
- ✅ 纯文件读取，性能高

### 写入：使用 VSTO ✅

**实现位置**：`Api/PowerPoint/PowerPointWriter.cs`

```csharp
// 1. 从模板文件打开
_app = new Application();
_presentation = _app.Presentations.Open(templatePath, ...);

// 2. 写入内容
_slideWriter.WriteSlides(jsonData.ContentSlides);

// 3. 另存为
_presentation.SaveAs(outputPath, ...);
```

**验证**：
- ✅ 使用 `Microsoft.Office.Interop.PowerPoint.Application`
- ✅ 从模板文件打开（`OpenFromTemplate`）
- ✅ 写入内容后另存为（`SaveAs`）
- ✅ 需要 PowerPoint 应用程序运行

## 🧪 运行测试

### 更新测试项目

测试项目目前使用 `.NET 8.0`，需要更新到 `.NET Framework 4.8` 以匹配主项目：

1. **编辑测试项目文件**
   - 文件：`OfficeHelperOpenxmVsto.Test\OfficeHelperOpenXml.Test.csproj`
   - 将 `<TargetFramework>net8.0</TargetFramework>` 改为 `<TargetFramework>net48</TargetFramework>`

2. **在 Visual Studio 中运行测试**
   - 打开测试资源管理器（测试 → 测试资源管理器）
   - 运行所有测试

### 手动测试

如果测试项目尚未更新，可以手动运行主程序进行测试：

```powershell
# 运行主程序
cd OfficeHelperOpenxmVsto\bin\Release\net48
.\OfficeHelperOpenXml.exe --help
```

## 📋 验证清单

- [ ] 在 Visual Studio 中成功构建项目
- [ ] 确认输出目录包含所有必要的 DLL
- [ ] 验证 COM 引用正确（Microsoft.Office.Core, Microsoft.Office.Interop.PowerPoint）
- [ ] 测试读取功能（使用 OpenXML SDK）
- [ ] 测试写入功能（使用 VSTO，需要 PowerPoint 安装）
- [ ] 运行现有测试确保功能正常

## 🔍 常见问题

### Q: 为什么不能使用 `dotnet build`？

**A:** .NET Core 版本的 MSBuild 不支持 COM 引用（`ResolveComReference`）。必须使用 Visual Studio 的 MSBuild（.NET Framework 版本）。

### Q: 如何确认 COM 引用正确？

**A:** 在 Visual Studio 中：
1. 右键点击项目 → **属性**
2. 选择 **引用**
3. 确认看到：
   - ✅ `Microsoft.Office.Core`
   - ✅ `Microsoft.Office.Interop.PowerPoint`
   - 这些引用应该显示为 **COM 引用**，而不是 NuGet 包

### Q: 构建时提示找不到 Office 互操作程序集？

**A:** 确保已安装 Microsoft Office（2016 或更高版本）。COM 引用需要系统安装的 Office PIA（Primary Interop Assemblies）。

## 📚 相关文档

- [VSTO 快速开始指南](VSTO_QUICK_START.md)
- [VSTO 迁移指南](VSTO_MIGRATION_GUIDE.md)
- [VSTO/COM/tlbimp 区别说明](VSTO_COM_TLBIMP_DIFFERENCES.md)

---

**注意**：由于 COM 引用的限制，建议始终在 Visual Studio 中开发和构建此项目。









