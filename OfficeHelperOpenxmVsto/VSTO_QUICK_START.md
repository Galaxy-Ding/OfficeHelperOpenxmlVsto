# VSTO 快速开始指南

## ✅ 迁移已完成

项目已成功迁移到 **VSTO 方式**，使用 **.NET Framework 4.8** 和 **COM 引用**。

## 🚀 立即开始

### 1. 在 Visual Studio 中打开项目

```powershell
# 打开解决方案
start OfficeHelperOpenxmVsto.sln
```

### 2. 验证 COM 引用

在 Visual Studio 中：
1. 右键点击 `OfficeHelperOpenXml` 项目
2. 选择 "属性" → "引用"
3. 确认看到：
   - ✅ `Microsoft.Office.Core`
   - ✅ `Microsoft.Office.Interop.PowerPoint`

### 3. 如果引用缺失

1. 右键点击项目 → **添加** → **引用**
2. 选择 **COM** 选项卡
3. 勾选：
   - `Microsoft Office 16.0 Object Library`
   - `Microsoft PowerPoint 16.0 Object Library`
4. 点击 **确定**

### 4. 构建项目

```powershell
# 在 Visual Studio 中
# 生成 → 生成解决方案 (Ctrl+Shift+B)

# 或使用命令行
msbuild OfficeHelperOpenXml.csproj /p:Configuration=Release
```

## 📋 主要变更

| 项目 | 之前 (.NET 8.0) | 现在 (VSTO) |
|------|----------------|------------|
| 目标框架 | net8.0 | net48 (.NET Framework 4.8) |
| Office 引用 | NuGet 包 + tlbimp | COM 引用（自动） |
| System.Drawing | NuGet 包 | 系统自带 |
| 互操作程序集 | 手动生成 | Visual Studio 自动生成 |

## ✨ VSTO 方式的优势

- ✅ **无需 tlbimp**：Visual Studio 自动处理
- ✅ **无需 NuGet 包**：使用系统安装的 Office PIA
- ✅ **更好的 IntelliSense**：完整的代码提示
- ✅ **标准化开发**：符合 Microsoft 官方推荐

## 🔧 开发环境要求

- ✅ Visual Studio 2019/2022（带 Office 开发工具）
- ✅ .NET Framework 4.8
- ✅ Microsoft Office 2016 或更高版本

## 📚 详细文档

- [完整迁移指南](VSTO_MIGRATION_GUIDE.md)
- [VSTO/COM/tlbimp 区别说明](VSTO_COM_TLBIMP_DIFFERENCES.md)

## ⚠️ 注意事项

1. **不再需要** `COM_INTEROP_SETUP.md` 中的 tlbimp 步骤
2. **不再需要** `Interop` 目录中的手动生成的 DLL
3. 所有互操作程序集由 Visual Studio 自动管理

## 🎯 下一步

1. 构建项目验证配置
2. 运行现有测试确保功能正常
3. 开始使用 VSTO 方式开发新功能

---

**迁移完成时间**：{{ 当前日期 }}
**目标框架**：.NET Framework 4.8
**开发方式**：VSTO (Visual Studio Tools for Office)









