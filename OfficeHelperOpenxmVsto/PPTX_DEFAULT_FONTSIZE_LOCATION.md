# PPTX 文件中默认字体大小的位置说明

## 概述

在 PowerPoint (PPTX) 文件中，默认字体大小可以在多个位置定义。本文档说明这些位置以及如何查看它们。

## 代码中的默认值

根据 `TextComponent.cs` 的实现，当无法从其他来源获取字体大小时，代码使用 **18pt** 作为默认值。

```294:294:OfficeHelperOpenxmVsto/Components/TextComponent.cs
                            runInfo.FontSize = 18f;
```

## PPTX 文件中默认字体大小的位置

### 1. 母版幻灯片 (Slide Master) - 最高优先级

**位置**: `ppt/slideMasters/slideMaster1.xml`

**路径结构**:
```
slideMaster
  └── textStyles
      ├── bodyStyle (正文样式)
      │   ├── lvl1PPr (Level 1 段落属性)
      │   │   └── defRPr (默认运行属性)
      │   │       └── sz (字体大小，单位：百分之一磅)
      │   ├── lvl2PPr (Level 2 段落属性)
      │   └── ...
      └── otherStyle (其他样式)
```

**XML 示例**:
```xml
<p:slideMaster>
  <p:textStyles>
    <p:bodyStyle>
      <a:lvl1PPr>
        <a:defRPr sz="1800"/>  <!-- 1800 = 18pt (1800/100) -->
      </a:lvl1PPr>
      <a:lvl2PPr>
        <a:defRPr sz="1600"/>  <!-- 1600 = 16pt -->
      </a:lvl2PPr>
    </p:bodyStyle>
  </p:textStyles>
</p:slideMaster>
```

**代码对应位置**:
```1115:1165:OfficeHelperOpenxmVsto/Components/TextComponent.cs
                // 1. 首先尝试从SlideMaster的TextStyles中获取（BodyStyle）
                var slideMasterPart = slidePart.SlideLayoutPart?.SlideMasterPart;
                Console.WriteLine($"[样式字体大小调试] 开始从样式获取字体大小，段落级别: {paragraphLevel}");
                
                if (slideMasterPart?.SlideMaster?.TextStyles != null)
                {
                    var textStyles = slideMasterPart.SlideMaster.TextStyles;
                    Console.WriteLine($"[样式字体大小调试] 找到SlideMaster的TextStyles");
                    
                    // 优先使用BodyStyle（正文样式），如果没有则使用OtherStyle
                    var bodyStyle = textStyles.BodyStyle;
                    if (bodyStyle != null)
                    {
                        Console.WriteLine($"[样式字体大小调试] 尝试从BodyStyle获取字体大小");
                        var fontSize = GetFontSizeFromTextStyleLevels(bodyStyle as OpenXmlCompositeElement, paragraphLevel);
                        if (fontSize.HasValue)
                        {
                            Console.WriteLine($"[样式字体大小调试] ✓ 从BodyStyle获取到字体大小: {fontSize.Value}pt (段落级别: {paragraphLevel})");
                            return fontSize.Value;
                        }
                        else
                        {
                            Console.WriteLine($"[样式字体大小调试] ✗ BodyStyle中没有找到字体大小 (段落级别: {paragraphLevel})");
                        }
                    }
                    else
                    {
                        Console.WriteLine($"[样式字体大小调试] BodyStyle为null");
                    }
                    
                    // 如果BodyStyle没有，尝试OtherStyle
                    var otherStyle = textStyles.OtherStyle;
                    if (otherStyle != null)
                    {
                        Console.WriteLine($"[样式字体大小调试] 尝试从OtherStyle获取字体大小");
                        var fontSize = GetFontSizeFromTextStyleLevels(otherStyle as OpenXmlCompositeElement, paragraphLevel);
                        if (fontSize.HasValue)
                        {
                            Console.WriteLine($"[样式字体大小调试] ✓ 从OtherStyle获取到字体大小: {fontSize.Value}pt (段落级别: {paragraphLevel})");
                            return fontSize.Value;
                        }
                        else
                        {
                            Console.WriteLine($"[样式字体大小调试] ✗ OtherStyle中没有找到字体大小 (段落级别: {paragraphLevel})");
                        }
                    }
                    else
                    {
                        Console.WriteLine($"[样式字体大小调试] OtherStyle为null");
                    }
                }
```

### 2. 演示文稿默认样式 (Presentation Default Text Style) - 次优先级

**位置**: `ppt/presentation.xml`

**路径结构**:
```
presentation
  └── defaultTextStyle
      ├── lvl1PPr (Level 1 段落属性)
      │   └── defRPr (默认运行属性)
      │       └── sz (字体大小)
      └── ...
```

**XML 示例**:
```xml
<p:presentation>
  <p:defaultTextStyle>
    <a:lvl1PPr>
      <a:defRPr sz="1800"/>  <!-- 1800 = 18pt -->
    </a:lvl1PPr>
  </p:defaultTextStyle>
</p:presentation>
```

**代码对应位置**:
```1171:1203:OfficeHelperOpenxmVsto/Components/TextComponent.cs
                // 2. 从Presentation的DefaultTextStyle中获取
                PresentationPart presentationPart = null;
                try
                {
                    // 通过SlidePart获取PresentationPart
                    presentationPart = slidePart.GetParentParts()
                        .OfType<PresentationPart>()
                        .FirstOrDefault();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[样式字体大小调试] 获取PresentationPart时出错: {ex.Message}");
                }
                
                if (presentationPart?.Presentation?.DefaultTextStyle != null)
                {
                    Console.WriteLine($"[样式字体大小调试] 尝试从Presentation的DefaultTextStyle获取字体大小");
                    var defaultTextStyle = presentationPart.Presentation.DefaultTextStyle;
                    var fontSize = GetFontSizeFromDefaultTextStyle(defaultTextStyle as OpenXmlCompositeElement, paragraphLevel);
                    if (fontSize.HasValue)
                    {
                        Console.WriteLine($"[样式字体大小调试] ✓ 从DefaultTextStyle获取到字体大小: {fontSize.Value}pt (段落级别: {paragraphLevel})");
                        return fontSize.Value;
                    }
                    else
                    {
                        Console.WriteLine($"[样式字体大小调试] ✗ DefaultTextStyle中没有找到字体大小 (段落级别: {paragraphLevel})");
                    }
                }
                else
                {
                    Console.WriteLine($"[样式字体大小调试] Presentation的DefaultTextStyle为null (presentationPart={(presentationPart != null ? "存在" : "null")})");
                }
```

### 3. 文本框架默认属性 (Text Body Default Run Properties)

**位置**: 在形状的 `TextBody` 元素中

**路径结构**:
```
shape
  └── txBody (文本体)
      ├── bodyPr (BodyProperties)
      │   └── defRPr (默认运行属性)
      │       └── sz (字体大小)
      └── p (段落)
```

**代码对应位置**:
```48:96:OfficeHelperOpenxmVsto/Components/TextComponent.cs
                // 获取文本框架的默认运行属性（用于字体大小继承）
                // 可能在BodyProperties中，也可能在TextBody的直接子元素中
                A.DefaultRunProperties textBodyDefaultRunProps = null;
                var bodyProps = textBody.BodyProperties;
                if (bodyProps != null)
                {
                    textBodyDefaultRunProps = bodyProps.GetFirstChild<A.DefaultRunProperties>();
                    if (textBodyDefaultRunProps != null)
                    {
                        Console.WriteLine($"[文本框架属性调试] 从BodyProperties获取到DefaultRunProperties");
                        if (textBodyDefaultRunProps.FontSize != null && textBodyDefaultRunProps.FontSize.HasValue)
                        {
                            Console.WriteLine($"[文本框架属性调试] 文本框架默认FontSize: {textBodyDefaultRunProps.FontSize.Value} (百分之一磅) = {textBodyDefaultRunProps.FontSize.Value / 100f}pt");
                        }
                        else
                        {
                            Console.WriteLine($"[文本框架属性调试] 文本框架DefaultRunProperties没有FontSize");
                        }
                    }
                    else
                    {
                        Console.WriteLine($"[文本框架属性调试] BodyProperties中没有DefaultRunProperties");
                    }
                }
                else
                {
                    Console.WriteLine($"[文本框架属性调试] BodyProperties为null");
                }
                // 如果BodyProperties中没有，尝试从TextBody的直接子元素中获取
                if (textBodyDefaultRunProps == null)
                {
                    textBodyDefaultRunProps = textBody.GetFirstChild<A.DefaultRunProperties>();
                    if (textBodyDefaultRunProps != null)
                    {
                        Console.WriteLine($"[文本框架属性调试] 从TextBody直接子元素获取到DefaultRunProperties");
                        if (textBodyDefaultRunProps.FontSize != null && textBodyDefaultRunProps.FontSize.HasValue)
                        {
                            Console.WriteLine($"[文本框架属性调试] 文本框架默认FontSize: {textBodyDefaultRunProps.FontSize.Value} (百分之一磅) = {textBodyDefaultRunProps.FontSize.Value / 100f}pt");
                        }
                        else
                        {
                            Console.WriteLine($"[文本框架属性调试] 文本框架DefaultRunProperties没有FontSize");
                        }
                    }
                    else
                    {
                        Console.WriteLine($"[文本框架属性调试] TextBody直接子元素中也没有DefaultRunProperties");
                    }
                }
```

### 4. 段落默认属性 (Paragraph Default Run Properties)

**位置**: 在段落的 `pPr` (段落属性) 中

**路径结构**:
```
p (段落)
  └── pPr (段落属性)
      └── defRPr (默认运行属性)
          └── sz (字体大小)
```

## 字体大小的优先级顺序

根据代码实现，字体大小的读取优先级如下：

1. **Run 元素的 FontSize** (最高优先级)
   - 直接在 `<a:rPr sz="..."/>` 中定义

2. **段落默认属性** (`paraDefaultRunProps`)
   - 在段落的 `pPr/defRPr` 中定义

3. **文本框架默认属性** (`textBodyDefaultRunProps`)
   - 在 `TextBody/bodyPr/defRPr` 或 `TextBody/defRPr` 中定义

4. **默认值 18pt** (当 `paragraphLevel = 0` 时直接使用，不从样式继承)
   - 代码硬编码的默认值

5. **母版样式** (仅当 `paragraphLevel > 0` 时)
   - `SlideMaster/TextStyles/BodyStyle` 或 `OtherStyle`

6. **演示文稿默认样式** (仅当 `paragraphLevel > 0` 时)
   - `Presentation/DefaultTextStyle`

**重要**: 当 `paragraphLevel = 0` 时，代码直接使用 18pt 默认值，**不会**从样式继承。

```290:311:OfficeHelperOpenxmVsto/Components/TextComponent.cs
                        // 调整优先级：先使用默认值18pt，如果段落级别>0，再从样式继承
                        if (paragraphLevel == 0)
                        {
                            // 段落级别为0时，直接使用默认值18pt，不从样式继承
                            runInfo.FontSize = 18f;
                            Console.WriteLine($"[字体大小调试] ✓ 段落级别为0，使用默认值FontSize: {runInfo.FontSize}pt (不从样式继承)");
                        }
                        else
                        {
                            // 段落级别>0时，先尝试从样式继承，如果没有则使用默认值18pt
                            float? styleFontSize = GetDefaultFontSizeFromStyles(slidePart, paragraphLevel);
                            if (styleFontSize.HasValue)
                            {
                                runInfo.FontSize = styleFontSize.Value;
                                Console.WriteLine($"[字体大小调试] ✓ 使用样式中的FontSize: {runInfo.FontSize}pt (段落级别: {paragraphLevel})");
                            }
                            else
                            {
                                runInfo.FontSize = 18f;
                                Console.WriteLine($"[字体大小调试] ✓ 样式中没有找到，使用默认值FontSize: {runInfo.FontSize}pt");
                            }
                        }
```

## 如何查看 PPTX 文件中的默认字体大小

### 方法 1: 使用解压工具查看 XML

1. 将 `.pptx` 文件重命名为 `.zip`
2. 解压 ZIP 文件
3. 导航到以下路径查看：
   - `ppt/slideMasters/slideMaster1.xml` - 查看母版样式
   - `ppt/presentation.xml` - 查看演示文稿默认样式

### 方法 2: 使用代码查看

运行程序时，查看控制台输出的调试信息：
- `[文本框架属性调试]` - 显示文本框架默认属性
- `[字体大小调试]` - 显示字体大小的读取过程
- `[样式字体大小调试]` - 显示从样式读取字体大小的过程
- `[TextStyle级别调试]` - 显示从不同级别读取字体大小的过程

### 方法 3: 使用 PowerPoint 查看

1. 打开 PowerPoint
2. 进入"视图" → "幻灯片母版"
3. 查看母版中的文本样式设置
4. 右键点击文本框 → "字体" → 查看默认字体大小

## 字体大小单位说明

- **XML 中的单位**: 百分之一磅 (1/100 point)
  - 例如：`sz="1800"` 表示 18pt (1800/100 = 18)
  
- **代码中的单位**: 磅 (point)
  - 代码会自动转换：`fontSize = xmlValue / 100f`

## 常见默认值

- **PowerPoint 2016/2019/365 默认**: 通常为 18pt (1800)
- **代码硬编码默认值**: 18pt
- **不同级别可能不同**:
  - Level 1: 通常 18pt
  - Level 2: 通常 16pt
  - Level 3: 通常 14pt
  - 等等...

## 注意事项

1. **段落级别为 0 的特殊处理**: 当 `paragraphLevel = 0` 时，代码直接使用 18pt，不从样式继承。这是代码的硬编码逻辑。

2. **样式继承的级别映射**:
   - `paragraphLevel = 0` → 尝试 Level0 或 Level1
   - `paragraphLevel = 1` → 使用 Level2
   - `paragraphLevel = 2` → 使用 Level3
   - 以此类推...

3. **如果所有来源都没有字体大小**: 最终使用 18pt 作为默认值。


