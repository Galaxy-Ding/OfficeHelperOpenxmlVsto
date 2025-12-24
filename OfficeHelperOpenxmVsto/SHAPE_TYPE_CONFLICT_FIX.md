# Shape 类型冲突修复说明

## 问题描述

出现以下错误：
```
"Shape"是"Microsoft.Office.Core.Shape"和"Microsoft.Office.Interop.PowerPoint.Shape"之间的不明确的引用
```

## 原因

代码中同时使用了：
- `using Microsoft.Office.Interop.PowerPoint;`
- `using Microsoft.Office.Core;`

这两个命名空间都包含 `Shape` 类型，导致编译器无法确定使用哪个。

## 解决方案

使用 **using 别名**来明确指定使用哪个 `Shape` 类型：

```csharp
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
// 使用别名避免 Shape 类型冲突
using PptShape = Microsoft.Office.Interop.PowerPoint.Shape;
```

然后在代码中使用 `PptShape` 而不是 `Shape`：

```csharp
// 之前（有冲突）
public Shape CreateShape(...)

// 之后（已修复）
public PptShape CreateShape(...)
```

## 已修复的文件

1. ✅ `Core/Writers/VstoShapeWriter.cs`
   - 所有 `Shape` 类型改为 `PptShape`
   - 所有方法签名和变量声明已更新

2. ✅ `Core/Writers/VstoStyleWriter.cs`
   - 所有方法参数中的 `Shape` 类型改为 `PptShape`
   - 移除了不支持的 `Strikethrough` 属性（已注释）

3. ✅ `OfficeHelperOpenXmlVsto.csproj`
   - 移除了重复的 COM 引用

## 剩余问题

如果仍然看到以下错误：
```
类型"MsoTriState"在未引用的程序集中定义。必须添加对程序集"office, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c"的引用。
```

这表示 COM 引用没有正确解析。请执行以下步骤：

### 在 Visual Studio 中修复 COM 引用

1. **打开 Visual Studio**
2. **打开项目** `OfficeHelperOpenXmlVsto.csproj`
3. **删除现有 COM 引用**（如果存在）
   - 在解决方案资源管理器中，展开项目
   - 展开 `引用` 节点
   - 如果看到 `Microsoft.Office.Core` 或 `Microsoft.Office.Interop.PowerPoint`，右键删除
4. **重新添加 COM 引用**
   - 右键点击 `引用` → `添加引用...`
   - 选择 **COM** 选项卡
   - 勾选：
     - ✅ `Microsoft Office 16.0 Object Library` (Microsoft.Office.Core)
     - ✅ `Microsoft PowerPoint 16.0 Object Library` (Microsoft.Office.Interop.PowerPoint)
   - 点击 **确定**
5. **清理并重新构建**
   - 生成 → 清理解决方案
   - 生成 → 重新生成解决方案

### 验证修复

构建成功后，应该不再看到：
- ❌ `"Shape"是"Microsoft.Office.Core.Shape"和"Microsoft.Office.Interop.PowerPoint.Shape"之间的不明确的引用`
- ❌ `类型"MsoTriState"在未引用的程序集中定义`

## 技术说明

### 为什么需要两个命名空间？

- `Microsoft.Office.Core`: 包含 Office 应用程序共用的类型（如 `MsoTriState`, `MsoAutoShapeType` 等）
- `Microsoft.Office.Interop.PowerPoint`: 包含 PowerPoint 特定的类型（如 `Shape`, `Slide`, `Presentation` 等）

### 为什么使用别名而不是完整命名空间？

使用别名 `PptShape` 比每次都写 `Microsoft.Office.Interop.PowerPoint.Shape` 更简洁，同时避免了类型冲突。

### 为什么不移除 `using Microsoft.Office.Core`？

因为代码中需要使用 `MsoTriState`, `MsoAutoShapeType` 等类型，这些类型在 `Microsoft.Office.Core` 命名空间中。如果移除这个 using，就需要在每个使用处都写完整命名空间，代码会变得冗长。

## 相关文档

- [COM 引用修复指南](COM_REFERENCE_FIX.md)
- [构建说明](BUILD_INSTRUCTIONS.md)

