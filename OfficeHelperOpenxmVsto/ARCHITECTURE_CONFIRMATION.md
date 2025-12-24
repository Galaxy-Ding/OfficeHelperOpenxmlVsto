# 架构确认文档

## ✅ 架构已确认

项目采用 **混合架构**：读取使用 OpenXML SDK，写入使用 VSTO。

---

## 📖 读取：OpenXML SDK

### 实现位置
- **主要类**：`Core/Readers/PresentationReader.cs`
- **API 接口**：`Api/PowerPointReader.cs`
- **工厂类**：`Api/PowerPointReaderFactory.cs`

### 代码示例

```csharp
// PresentationReader.cs
public PresentationInfo ReadPresentation(string filePath)
{
    using (var doc = PresentationDocument.Open(filePath, false))
    {
        var presentationPart = doc.PresentationPart;
        // 读取幻灯片、样式等信息
    }
}
```

### 特点
- ✅ **无需 PowerPoint 应用程序**：纯文件读取
- ✅ **高性能**：直接解析文件格式
- ✅ **跨平台**：可在没有 Office 的环境中运行
- ✅ **依赖**：`DocumentFormat.OpenXml` NuGet 包

### 使用场景
- 分析 PowerPoint 文件结构
- 提取文本、图片、形状等信息
- 生成 JSON 报告
- 文件验证和比较

---

## ✍️ 写入：VSTO (Visual Studio Tools for Office)

### 实现位置
- **主要类**：`Api/PowerPoint/PowerPointWriter.cs`
- **API 接口**：`Api/PowerPoint/IPowerPointWriter.cs`
- **工厂类**：`Api/PowerPoint/PowerPointWriterFactory.cs`
- **VSTO 写入器**：`Core/Writers/VstoSlideWriter.cs`

### 代码示例

```csharp
// PowerPointWriter.cs
public bool OpenFromTemplate(string templatePath)
{
    // 1. 启动 PowerPoint 应用程序
    _app = new Application();
    _app.Visible = MsoTriState.msoFalse;
    
    // 2. 从模板文件打开
    _presentation = _app.Presentations.Open(
        templatePath,
        WithWindow: MsoTriState.msoFalse,
        ReadOnly: MsoTriState.msoTrue
    );
    
    // 3. 初始化写入器
    _slideWriter = new VstoSlideWriter(_presentation);
    return true;
}

public bool WriteFromJson(string jsonData)
{
    // 写入内容幻灯片
    _slideWriter.WriteSlides(jsonData.ContentSlides);
    return true;
}

public bool SaveAs(string outputPath)
{
    // 另存为
    _presentation.SaveAs(outputPath, PpSaveAsFileType.ppSaveAsDefault);
    return true;
}
```

### 工作流程

```
1. 打开模板文件 (OpenFromTemplate)
   ↓
2. 清空内容幻灯片 (ClearAllContentSlides) [可选]
   ↓
3. 写入内容 (WriteFromJson)
   ↓
4. 另存为 (SaveAs)
   ↓
5. 关闭和清理 (Dispose)
```

### 特点
- ✅ **完整功能支持**：支持所有 PowerPoint 功能
- ✅ **格式保真**：保持模板的格式和样式
- ✅ **需要 PowerPoint**：需要安装 Microsoft Office
- ✅ **依赖**：COM 引用（Microsoft.Office.Interop.PowerPoint）

### 使用场景
- 从 JSON 数据生成 PowerPoint 文件
- 基于模板创建演示文稿
- 批量生成幻灯片
- 需要复杂格式和动画的场景

---

## 🔄 完整工作流程

### 典型使用场景

```csharp
// 1. 读取（使用 OpenXML SDK）
using (var reader = PowerPointReaderFactory.CreateReader(templatePath, out bool success))
{
    if (success)
    {
        string json = reader.ToJson();
        // 分析或修改 JSON 数据
    }
}

// 2. 写入（使用 VSTO）
using (var writer = PowerPointWriterFactory.CreateWriter())
{
    writer.OpenFromTemplate(templatePath);
    writer.ClearAllContentSlides();
    writer.WriteFromJson(modifiedJson);
    writer.SaveAs(outputPath);
}
```

### 便捷方法

```csharp
// OfficeHelperWrapper.cs
public static bool WritePowerPointFromJson(
    string templatePath, 
    string jsonData, 
    string outputPath)
{
    // 内部使用 VSTO 方式
    using (var writer = PowerPointWriterFactory.CreateWriter())
    {
        return writer.OpenFromTemplate(templatePath) &&
               writer.ClearAllContentSlides() &&
               writer.WriteFromJson(jsonData) &&
               writer.SaveAs(outputPath);
    }
}
```

---

## 📊 架构对比

| 特性 | OpenXML SDK (读取) | VSTO (写入) |
|------|-------------------|------------|
| **用途** | 读取和分析 | 写入和生成 |
| **需要 Office** | ❌ 不需要 | ✅ 需要 |
| **性能** | ⚡ 快速 | 🐢 较慢（需要启动应用） |
| **功能完整性** | ⚠️ 有限 | ✅ 完整 |
| **格式保真** | ⚠️ 可能丢失 | ✅ 完美保持 |
| **跨平台** | ✅ 是 | ❌ 否（Windows + Office） |
| **依赖** | NuGet 包 | COM 引用 |

---

## ✅ 验证清单

### 读取功能（OpenXML SDK）
- [x] `PresentationReader` 使用 `PresentationDocument.Open`
- [x] 无需 PowerPoint 应用程序
- [x] 可以提取所有元素信息
- [x] 生成 JSON 输出

### 写入功能（VSTO）
- [x] `PowerPointWriter` 使用 `Application.Presentations.Open`
- [x] 从模板文件打开
- [x] 写入内容后另存为
- [x] 需要 PowerPoint 应用程序运行
- [x] 正确释放 COM 对象

---

## 🎯 下一步

1. **构建项目**（在 Visual Studio 中）
   - 参考 [构建和测试指南](BUILD_AND_TEST.md)

2. **运行测试**
   - 更新测试项目到 net48 ✅
   - 在 Visual Studio 中运行测试

3. **开发新功能**
   - 读取功能：使用 OpenXML SDK
   - 写入功能：使用 VSTO 方式

---

## 📚 相关文档

- [构建和测试指南](BUILD_AND_TEST.md)
- [VSTO 快速开始指南](VSTO_QUICK_START.md)
- [VSTO 迁移指南](VSTO_MIGRATION_GUIDE.md)

---

**最后更新**：2025-12-19
**架构状态**：✅ 已确认并实现

