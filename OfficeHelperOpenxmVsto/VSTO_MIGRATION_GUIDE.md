# VSTO 迁移指南

## 概述

本项目已从 .NET 8.0 迁移到 **.NET Framework 4.8**，以支持完整的 **VSTO (Visual Studio Tools for Office)** 功能。

## 主要变更

### 1. 目标框架变更

**之前**：`.NET 8.0`
```xml
<TargetFramework>net8.0</TargetFramework>
```

**现在**：`.NET Framework 4.8`
```xml
<TargetFramework>net48</TargetFramework>
```

### 2. COM 引用方式变更

**之前**（.NET 8.0）：
- 使用 NuGet 包：`Microsoft.Office.Interop.PowerPoint`
- 使用 tlbimp 手动生成：`Microsoft.Office.Core.dll`

**现在**（.NET Framework 4.8 + VSTO）：
- 使用 **COM 引用**（Visual Studio 自动处理）
- 无需手动运行 tlbimp
- 无需 NuGet 包（使用系统安装的 Office PIA）

### 3. 项目文件配置

#### 移除的配置
- ❌ `EnableComHosting`（.NET 8.0 特有）
- ❌ `Microsoft.Office.Interop.PowerPoint` NuGet 包
- ❌ `System.Drawing.Common` NuGet 包（.NET Framework 自带）
- ❌ 手动 tlbimp 生成逻辑

#### 新增的配置
- ✅ `COMReference` 元素（VSTO 标准方式）
- ✅ `Microsoft.Office.Core` COM 引用
- ✅ `Microsoft.Office.Interop.PowerPoint` COM 引用

## 使用 VSTO 方式的优势

### 1. 简化的开发体验
- ✅ Visual Studio 自动生成互操作程序集
- ✅ 无需手动运行 tlbimp
- ✅ 更好的 IntelliSense 支持
- ✅ 自动处理 COM 对象生命周期

### 2. 更好的兼容性
- ✅ 与 Office 版本完全匹配
- ✅ 使用系统安装的 PIA（Primary Interop Assemblies）
- ✅ 减少版本冲突问题

### 3. 标准化的开发方式
- ✅ 符合 Microsoft 官方推荐方式
- ✅ 与其他 VSTO 项目保持一致
- ✅ 便于团队协作

## 开发环境要求

### 必需组件

1. **Visual Studio 2019 或 2022**
   - 安装时选择 "Office/SharePoint 开发" 工作负载
   - 或单独安装 "Visual Studio Tools for Office"

2. **.NET Framework 4.8**
   - 通常随 Visual Studio 自动安装
   - 或从 [Microsoft 官网](https://dotnet.microsoft.com/download/dotnet-framework/net48) 下载

3. **Microsoft Office**
   - 建议安装 Office 2016 或更高版本
   - 确保安装了 PowerPoint

### 验证安装

在 Visual Studio 中：
1. 创建新项目
2. 查看项目模板中是否有 "Office Add-in" 选项
3. 如果有，说明 VSTO 已正确安装

## 首次设置步骤

### 方法 1：使用 Visual Studio（推荐）

1. **打开项目**
   ```powershell
   # 在 Visual Studio 中打开解决方案
   OfficeHelperOpenxmVsto.sln
   ```

2. **验证 COM 引用**
   - 右键点击项目 → 属性
   - 查看 "引用" 部分
   - 应该看到：
     - `Microsoft.Office.Core`
     - `Microsoft.Office.Interop.PowerPoint`

3. **如果引用缺失**
   - 右键点击项目 → 添加 → 引用
   - 选择 "COM" 选项卡
   - 勾选：
     - `Microsoft Office 16.0 Object Library`
     - `Microsoft PowerPoint 16.0 Object Library`
   - 点击确定

4. **构建项目**
   ```powershell
   dotnet build
   # 或
   msbuild OfficeHelperOpenXml.csproj
   ```

### 方法 2：使用 MSBuild 命令行

```powershell
# 确保已安装 .NET Framework 4.8 Developer Pack
# 然后直接构建
msbuild OfficeHelperOpenXml.csproj /p:Configuration=Release
```

## COM 引用说明

### Microsoft.Office.Core
- **GUID**: `{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}`
- **版本**: 2.8
- **用途**: Office 核心对象模型（如 `MsoTriState`、`MsoShapeType` 等）

### Microsoft.Office.Interop.PowerPoint
- **GUID**: `{91493440-5A91-11CF-8700-00AA0060263B}`
- **版本**: 1.8
- **用途**: PowerPoint 应用程序对象模型

## 代码变更说明

### 无需修改的代码

大部分代码**无需修改**，因为：
- COM 互操作 API 保持不变
- 命名空间保持不变（`Microsoft.Office.Interop.PowerPoint`、`Microsoft.Office.Core`）
- 方法签名保持不变

### 可能需要调整的地方

1. **System.Drawing 引用**
   - .NET Framework 4.8 自带 `System.Drawing`
   - 如果之前使用了 `System.Drawing.Common` NuGet 包，需要改为直接引用 `System.Drawing`

2. **异步代码**
   - .NET Framework 4.8 的异步支持与 .NET 8.0 略有不同
   - 如果使用了最新的异步特性，可能需要调整

## 构建和部署

### 构建项目

```powershell
# 使用 MSBuild
msbuild OfficeHelperOpenXml.csproj /p:Configuration=Release

# 或使用 Visual Studio
# 生成 → 生成解决方案 (Ctrl+Shift+B)
```

### 输出文件

构建后，互操作程序集会自动复制到输出目录：
```
bin\Release\
├── OfficeHelperOpenXml.exe
├── Microsoft.Office.Interop.PowerPoint.dll
└── Microsoft.Office.Core.dll (如果需要)
```

### 部署要求

目标机器需要：
1. **.NET Framework 4.8** 运行时
2. **Microsoft Office**（与开发环境版本兼容）
3. **Visual C++ 运行时**（如果需要）

## 常见问题

### Q1: 构建时提示找不到 COM 组件

**解决方案**：
1. 确保已安装 Microsoft Office
2. 在 Visual Studio 中重新添加 COM 引用
3. 检查 Office 版本是否与 COM 引用版本匹配

### Q2: 运行时提示找不到互操作程序集

**解决方案**：
1. 确保互操作程序集已复制到输出目录
2. 检查 `EmbedInteropTypes` 设置（应设为 `False`）
3. 确保目标机器安装了相同版本的 Office

### Q3: 如何更新 Office 版本？

**解决方案**：
1. 在 Visual Studio 中移除旧的 COM 引用
2. 添加新版本的 COM 引用
3. 重新构建项目

### Q4: 可以在 .NET Core/.NET 8.0 中使用 VSTO 吗？

**答案**：不可以。VSTO 的完整功能（包括 COM 引用）仅支持 .NET Framework。如果必须在 .NET 8.0 中使用，需要使用：
- NuGet 包（如 `Microsoft.Office.Interop.PowerPoint`）
- 或手动使用 tlbimp 生成互操作程序集

## 迁移检查清单

- [x] 项目文件已更新为 .NET Framework 4.8
- [x] COM 引用已配置
- [x] 移除了不必要的 NuGet 包
- [ ] 验证代码可以正常编译
- [ ] 验证代码可以正常运行
- [ ] 测试 PowerPoint 互操作功能
- [ ] 更新文档和 README

## 下一步

1. **构建项目**验证配置是否正确
2. **运行测试**确保功能正常
3. **更新文档**反映新的配置方式
4. **团队通知**告知其他开发者新的开发环境要求

## 参考资源

- [VSTO 开发文档](https://docs.microsoft.com/visualstudio/vsto/)
- [Office 主互操作程序集](https://docs.microsoft.com/visualstudio/vsto/office-primary-interop-assemblies)
- [.NET Framework 4.8 下载](https://dotnet.microsoft.com/download/dotnet-framework/net48)

---

**注意**：迁移到 VSTO 方式后，不再需要 `COM_INTEROP_SETUP.md` 中的 tlbimp 步骤。所有互操作程序集由 Visual Studio 自动管理。









