# PowerPointWriter API 文档

## 概述

`PowerPointWriter` 是一个基于 VSTO (Visual Studio Tools for Office) 的 PowerPoint 写入器，用于从 JSON 数据生成 PowerPoint 演示文稿。它支持从模板文件创建新的演示文稿，保留模板的母版样式和格式。

## 核心功能

- ✅ 从模板文件打开演示文稿
- ✅ 清除内容幻灯片（保留母版）
- ✅ 从 JSON 数据写入内容
- ✅ 保存为新的 PPTX 文件
- ✅ 自动资源管理和 COM 对象释放

## 快速开始

### 基本使用

```csharp
using OfficeHelperOpenXml.Api.PowerPoint;
using OfficeHelperOpenXml.Models.Json;

// 创建写入器
using (var writer = PowerPointWriterFactory.CreateWriter())
{
    // 1. 打开模板文件
    if (!writer.OpenFromTemplate("template.pptx"))
    {
        Console.WriteLine("打开模板失败");
        return;
    }

    // 2. 清除内容幻灯片（可选）
    writer.ClearAllContentSlides();

    // 3. 准备 JSON 数据
    var jsonData = new PresentationJsonData
    {
        ContentSlides = new List<SlideJsonData>
        {
            new SlideJsonData
            {
                PageNumber = 1,
                Title = "第一张幻灯片",
                Shapes = new List<ShapeJsonData>
                {
                    new ShapeJsonData
                    {
                        Type = "textbox",
                        Name = "Title",
                        Box = "2,2,20,3",
                        HasText = 1,
                        Text = new List<TextRunJsonData>
                        {
                            new TextRunJsonData
                            {
                                Content = "Hello, World!",
                                Font = "Arial",
                                FontSize = 24,
                                FontColor = "RGB(0,0,0)"
                            }
                        }
                    }
                }
            }
        }
    };

    // 4. 写入数据
    writer.WriteFromJsonData(jsonData);

    // 5. 保存文件
    writer.SaveAs("output.pptx");
}
```

### 使用便捷方法

```csharp
using OfficeHelperOpenXml.Api;

// 使用便捷方法（推荐）
string templatePath = "26xdemo2.pptx";
string jsonData = File.ReadAllText("content.json");
string outputPath = "output.pptx";

bool success = OfficeHelperWrapper.WritePowerPointFromJson(
    templatePath, jsonData, outputPath);

if (success)
{
    Console.WriteLine("文件生成成功！");
}
```

## API 参考

### PowerPointWriterFactory

#### CreateWriter()

创建 PowerPoint 写入器实例。

```csharp
public static IPowerPointWriter CreateWriter()
```

**返回**: `IPowerPointWriter` 实例

**示例**:
```csharp
using (var writer = PowerPointWriterFactory.CreateWriter())
{
    // 使用 writer
}
```

---

### IPowerPointWriter 接口

#### OpenFromTemplate(string templatePath)

从模板文件打开演示文稿。

**参数**:
- `templatePath` (string): 模板文件路径

**返回**: `bool` - 是否成功

**异常处理**:
- 文件不存在：返回 `false`，记录错误日志
- PowerPoint 不可用：返回 `false`，记录错误日志
- COM 错误：捕获并记录详细错误信息

**示例**:
```csharp
bool success = writer.OpenFromTemplate("26xdemo2.pptx");
```

---

#### ClearAllContentSlides()

清除所有内容幻灯片的内容（保留母版形状和样式）。

**返回**: `bool` - 是否成功

**说明**:
- 只删除非母版形状
- 保留母版上的所有形状和样式
- 从后往前删除，避免索引问题

**示例**:
```csharp
bool success = writer.ClearAllContentSlides();
```

---

#### WriteFromJson(string jsonData)

从 JSON 字符串写入内容。

**参数**:
- `jsonData` (string): JSON 数据字符串

**返回**: `bool` - 是否成功

**示例**:
```csharp
string json = File.ReadAllText("data.json");
bool success = writer.WriteFromJson(json);
```

---

#### WriteFromJsonData(PresentationJsonData jsonData)

从 PresentationJsonData 对象写入内容。

**参数**:
- `jsonData` (PresentationJsonData): JSON 数据对象

**返回**: `bool` - 是否成功

**示例**:
```csharp
var jsonData = new PresentationJsonData { /* ... */ };
bool success = writer.WriteFromJsonData(jsonData);
```

---

#### SaveAs(string outputPath)

保存到文件。

**参数**:
- `outputPath` (string): 输出文件路径

**返回**: `bool` - 是否成功

**说明**:
- 自动创建输出目录（如果不存在）
- 如果文件已存在，会尝试覆盖
- 保存后验证文件是否存在

**示例**:
```csharp
bool success = writer.SaveAs("output.pptx");
```

---

#### Close()

关闭文档（不关闭 PowerPoint 应用程序）。

**说明**:
- 通常在 `Dispose()` 中自动调用
- 手动调用可用于提前关闭文档

---

#### Dispose()

释放资源。

**说明**:
- 实现 `IDisposable` 接口
- 自动关闭文档和 PowerPoint 应用程序
- 释放所有 COM 对象
- 强制垃圾回收

**示例**:
```csharp
using (var writer = PowerPointWriterFactory.CreateWriter())
{
    // 使用 writer
    // 自动调用 Dispose()
}
```

## JSON 数据格式

### PresentationJsonData

```json
{
  "content_slides": [
    {
      "page_number": 1,
      "title": "幻灯片标题",
      "shapes": [
        {
          "type": "textbox",
          "name": "Shape1",
          "box": "2,2,10,3",
          "has_text": 1,
          "text": [
            {
              "content": "文本内容",
              "font": "Arial",
              "font_size": 14,
              "font_color": "RGB(0,0,0)"
            }
          ],
          "fill": {
            "color": "RGB(255,255,255)",
            "opacity": 1.0
          },
          "line": {
            "has_outline": 1,
            "color": "RGB(0,0,0)",
            "width": 1.0
          }
        }
      ]
    }
  ]
}
```

### 支持的形状类型

- `textbox`: 文本框
- `autoshape`: 自动形状（矩形、圆形等）
- `picture`: 图片（占位符）
- `table`: 表格
- `group`: 组合形状
- `connection`: 连接线

## 模板处理流程

```
1. 检查模板文件存在
   ↓
2. 检查 PowerPoint 可用
   ↓
3. 启动 PowerPoint 应用程序（后台运行）
   ↓
4. 打开模板文件
   ↓
5. 清除内容幻灯片（可选）
   ↓
6. 解析 JSON 数据
   ↓
7. 创建/获取幻灯片
   ↓
8. 创建形状并应用样式
   ↓
9. 保存文件
   ↓
10. 关闭文档并释放资源
```

## 错误处理

### 常见错误

1. **模板文件不存在**
   - 错误信息：`模板文件不存在: {path}`
   - 解决方案：检查文件路径是否正确

2. **PowerPoint 不可用**
   - 错误信息：`PowerPoint 不可用，请确保已安装 Microsoft PowerPoint`
   - 解决方案：安装 Microsoft PowerPoint

3. **COM 错误**
   - 错误信息：`COM 错误：{message} (HRESULT: 0x{code:X})`
   - 解决方案：检查 PowerPoint 是否正常运行，重启应用程序

4. **文件保存失败**
   - 错误信息：`保存文件失败: {message}`
   - 解决方案：检查输出路径权限，确保有写入权限

### 日志记录

所有操作都会记录日志：
- ✅ 成功操作：使用 `LogSuccess()`
- ⚠️ 警告：使用 `LogWarning()`
- ❌ 错误：使用 `LogError()`

## 性能优化

1. **批量写入优化**
   - 写入多个幻灯片时，减少日志输出
   - 从后往前删除形状，避免索引问题

2. **资源管理**
   - 使用 `using` 语句确保资源释放
   - 自动释放 COM 对象
   - 强制垃圾回收

3. **错误恢复**
   - 单个形状失败不影响其他形状
   - 单个幻灯片失败不影响其他幻灯片

## 最佳实践

1. **始终使用 using 语句**
   ```csharp
   using (var writer = PowerPointWriterFactory.CreateWriter())
   {
       // 使用 writer
   }
   ```

2. **检查返回值**
   ```csharp
   if (!writer.OpenFromTemplate(templatePath))
   {
       // 处理错误
       return;
   }
   ```

3. **使用绝对路径**
   - 模板路径和输出路径建议使用绝对路径
   - 避免相对路径导致的路径问题

4. **处理异常**
   ```csharp
   try
   {
       writer.WriteFromJsonData(jsonData);
   }
   catch (Exception ex)
   {
       // 记录并处理异常
       Logger.LogError(ex.Message);
   }
   ```

## 限制和注意事项

1. **需要 PowerPoint**
   - 必须在安装了 Microsoft PowerPoint 的 Windows 系统上运行
   - 不支持跨平台

2. **COM 对象管理**
   - 必须正确释放 COM 对象，否则可能导致资源泄漏
   - 使用 `using` 语句可以自动管理

3. **文件锁定**
   - 如果模板文件正在被其他程序打开，可能无法打开
   - 输出文件如果已存在且被打开，可能无法覆盖

4. **性能**
   - 大量形状（100+）时写入可能较慢
   - 建议分批处理或优化 JSON 数据

## 示例代码

### 完整示例

```csharp
using System;
using System.Collections.Generic;
using OfficeHelperOpenXml.Api.PowerPoint;
using OfficeHelperOpenXml.Models.Json;

class Program
{
    static void Main()
    {
        string templatePath = "26xdemo2.pptx";
        string outputPath = "output.pptx";

        using (var writer = PowerPointWriterFactory.CreateWriter())
        {
            // 打开模板
            if (!writer.OpenFromTemplate(templatePath))
            {
                Console.WriteLine("打开模板失败");
                return;
            }

            // 清除内容
            writer.ClearAllContentSlides();

            // 创建数据
            var jsonData = CreateSampleData();

            // 写入数据
            if (!writer.WriteFromJsonData(jsonData))
            {
                Console.WriteLine("写入数据失败");
                return;
            }

            // 保存文件
            if (!writer.SaveAs(outputPath))
            {
                Console.WriteLine("保存文件失败");
                return;
            }

            Console.WriteLine("成功生成文件！");
        }
    }

    static PresentationJsonData CreateSampleData()
    {
        return new PresentationJsonData
        {
            ContentSlides = new List<SlideJsonData>
            {
                new SlideJsonData
                {
                    PageNumber = 1,
                    Title = "示例幻灯片",
                    Shapes = new List<ShapeJsonData>
                    {
                        new ShapeJsonData
                        {
                            Type = "textbox",
                            Name = "Title",
                            Box = "2,2,20,3",
                            HasText = 1,
                            Text = new List<TextRunJsonData>
                            {
                                new TextRunJsonData
                                {
                                    Content = "示例标题",
                                    Font = "Arial",
                                    FontSize = 24,
                                    FontColor = "RGB(0,0,0)",
                                    FontBold = 1
                                }
                            }
                        }
                    }
                }
            }
        };
    }
}
```

## 相关文档

- [架构确认文档](ARCHITECTURE_CONFIRMATION.md)
- [VSTO 快速开始指南](VSTO_QUICK_START.md)
- [JSON 格式说明](JSON_FORMAT.md)

---

**最后更新**: 2025-12-19

