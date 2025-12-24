# 文本框属性测试模板说明

## 模板文件位置

`textbox_properties_test_template.pptx`

## 模板要求

创建一个包含9个文本框的PPTX文件，每个文本框测试不同的属性组合。

### 文本框配置

#### 文本框1：基础属性
- **内容**: `Arial-(255,0,0)-14pt`
- **字体**: Arial
- **字号**: 14pt
- **颜色**: RGB(255, 0, 0) - 红色
- **属性**: 
  - ✓ 粗体 (Bold)
  - ✓ 斜体 (Italic)
  - ✓ 下划线 (Underline)
  - ✓ 删除线 (Strikethrough)

#### 文本框2：阴影属性
- **内容**: `Times New Roman-(0,0,255)-16pt`
- **字体**: Times New Roman
- **字号**: 16pt
- **颜色**: RGB(0, 0, 255) - 蓝色
- **属性**: 
  - ✓ 外阴影 (Outer Shadow)
  - 阴影颜色: RGB(128, 128, 128)
  - 模糊半径: 5pt
  - 距离: 3pt
  - 角度: 45度
  - 透明度: 50%

#### 文本框3：文本填充
- **内容**: `Calibri-(0,128,0)-18pt`
- **字体**: Calibri
- **字号**: 18pt
- **颜色**: RGB(0, 128, 0) - 绿色
- **属性**: 
  - ✓ 渐变填充 (Gradient Fill)
  - 渐变类型: 线性渐变
  - 渐变角度: 90度
  - 渐变停止点: 
    - 0%: RGB(0, 128, 0)
    - 100%: RGB(0, 255, 0)

#### 文本框4：文本轮廓
- **内容**: `Verdana-(255,128,0)-15pt`
- **字体**: Verdana
- **字号**: 15pt
- **颜色**: RGB(255, 128, 0) - 橙色
- **属性**: 
  - ✓ 文本轮廓 (Text Outline)
  - 轮廓宽度: 2pt
  - 轮廓颜色: RGB(0, 0, 0)
  - 虚线样式: 实线
  - 透明度: 0%

#### 文本框5：文本效果
- **内容**: `Courier New-(128,0,255)-17pt`
- **字体**: Courier New
- **字号**: 17pt
- **颜色**: RGB(128, 0, 255) - 紫色
- **属性**: 
  - ✓ 发光效果 (Glow)
    - 发光半径: 10pt
    - 发光颜色: RGB(255, 0, 255)
  - ✓ 反射效果 (Reflection)
    - 模糊半径: 2pt
    - 距离: 5pt
  - ✓ 柔边效果 (Soft Edge)
    - 柔边半径: 3pt

#### 文本框6：组合属性
- **内容**: `Microsoft YaHei-(255,0,128)-16pt`
- **字体**: Microsoft YaHei
- **字号**: 16pt
- **颜色**: RGB(255, 0, 128) - 粉红色
- **属性**: 
  - ✓ 粗体 + 斜体
  - ✓ 渐变填充
  - ✓ 文本轮廓
  - ✓ 阴影效果
  - ✓ 发光效果

#### 文本框7：主题颜色和颜色变换
- **内容**: `SimSun-(0,128,128)-14pt`
- **字体**: SimSun (宋体)
- **字号**: 14pt
- **颜色**: 主题颜色 - Accent1
- **属性**: 
  - ✓ 主题颜色 (Theme Color)
  - ✓ 颜色变换 (Color Transforms)
    - LumMod: 50000 (50%)
    - Tint: 50000 (50%)

#### 文本框8：高亮颜色
- **内容**: `Arial-(255,255,0)-15pt`
- **字体**: Arial
- **字号**: 15pt
- **颜色**: RGB(255, 255, 0) - 黄色
- **属性**: 
  - ✓ 高亮背景色 (Highlight Color)
  - 高亮颜色: RGB(255, 255, 0)

#### 文本框9：字符间距、上标/下标
- **内容**: `Calibri-(64,64,64)-13pt`
- **字体**: Calibri
- **字号**: 13pt
- **颜色**: RGB(64, 64, 64) - 灰色
- **属性**: 
  - ✓ 字符间距 (Character Spacing): 2pt
  - ✓ 上标 (Superscript): 部分文本
  - ✓ 下标 (Subscript): 部分文本

## 创建步骤

1. 打开 PowerPoint
2. 创建新演示文稿
3. 在第一张幻灯片上创建9个文本框
4. 按照上述配置设置每个文本框的属性
5. 保存为 `textbox_properties_test_template.pptx`
6. 将文件放置在 `OfficeHelperOpenxmVsto.Test/TestTemplates/` 目录下

## 验证

运行测试验证模板是否正确：

```bash
dotnet test --filter TextboxPropertyCompletenessTest
```

或使用 Program.cs 的测试入口：

```bash
dotnet run --project OfficeHelperOpenxmVsto.Test -- --textbox-props
```

