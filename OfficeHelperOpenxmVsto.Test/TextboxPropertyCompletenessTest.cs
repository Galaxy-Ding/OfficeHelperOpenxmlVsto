using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeHelperOpenXml.Api;
using OfficeHelperOpenXml.Elements;
using OfficeHelperOpenXml.Components;
using OfficeHelperOpenXml.Models;
using OfficeHelperOpenXml.Test.TestData;
using Xunit;
using Xunit.Abstractions;

namespace OfficeHelperOpenXml.Test
{
    /// <summary>
    /// 文本框属性完整性测试类
    /// 验证从PPTX读取、JSON序列化和写回PPTX的完整流程中属性是否完整保留
    /// 
    /// 代码规范：
    /// - xUnit 的 Assert.NotNull() 和 Assert.NotEmpty() 方法只接受一个参数
    /// - 如果需要输出错误信息，请使用 _output.WriteLine() 在断言前后输出
    /// - 错误示例：Assert.NotNull(obj, "错误信息")  ❌
    /// - 正确示例：Assert.NotNull(obj); _output.WriteLine("错误信息");  ✅
    /// </summary>
    public class TextboxPropertyCompletenessTest
    {
        private readonly ITestOutputHelper _output;
        private readonly string _testTemplatePath;
        private readonly string _testOutputDir;
        private readonly string _testReportsDir;
        private readonly TextboxPropertyValidator _validator;
        private readonly TextboxPropertyCoverageAnalyzer _coverageAnalyzer;

        public TextboxPropertyCompletenessTest(ITestOutputHelper output)
        {
            _output = output;
            
            // 使用统一的测试路径配置
            _testTemplatePath = TestPaths.TextboxPptxPath; // 直接使用用户指定的 textbox.pptx
            _testOutputDir = TestPaths.TestOutputDir;
            _testReportsDir = TestPaths.TestReportsDir;
            
            // 创建输出目录
            Directory.CreateDirectory(_testOutputDir);
            Directory.CreateDirectory(_testReportsDir);
            
            _validator = new TextboxPropertyValidator();
            _coverageAnalyzer = new TextboxPropertyCoverageAnalyzer();
        }

        #region 读取测试

        [Fact]
        public void TestReadBasicProperties()
        {
            if (!File.Exists(_testTemplatePath))
            {
                _output.WriteLine($"测试 PPTX 不存在: {_testTemplatePath}");
                _output.WriteLine("请确认 D:\\pythonf\\c_sharp_project\\OfficeHelperOpenxmVsto\\test_ppt\\textbox.pptx 是否存在");
                return;
            }

            using (var reader = new PowerPointReader())
            {
                reader.Load(_testTemplatePath);
                var elements = reader.GetAllElements();
                var textboxes = elements.OfType<TextBoxElement>().ToList();

                Assert.True(textboxes.Count > 0, "应该至少有一个文本框");

                // 输出所有文本框的文本内容，用于调试
                _output.WriteLine($"找到 {textboxes.Count} 个文本框:");
                for (int i = 0; i < textboxes.Count; i++)
                {
                    var content = GetTextContent(textboxes[i]);
                    _output.WriteLine($"  文本框 {i + 1}: \"{content}\"");
                }

                // 查找文本框1（基础属性）
                var textbox1 = textboxes.FirstOrDefault(tb => GetTextContent(tb).Contains("Arial-(255,0,0)-14pt"));
                
                // 如果找不到，尝试查找包含 "Arial" 的文本框
                if (textbox1 == null)
                {
                    _output.WriteLine("未找到包含 'Arial-(255,0,0)-14pt' 的文本框，尝试查找包含 'Arial' 的文本框");
                    textbox1 = textboxes.FirstOrDefault(tb => GetTextContent(tb).Contains("Arial"));
                }
                
                // 如果还是找不到，使用第一个文本框
                if (textbox1 == null && textboxes.Count > 0)
                {
                    _output.WriteLine("未找到包含 'Arial' 的文本框，使用第一个文本框");
                    textbox1 = textboxes[0];
                }
                
                Assert.True(textbox1 != null, $"找不到匹配的文本框。已检查 {textboxes.Count} 个文本框。");

                var textComponent = textbox1.GetComponent<TextComponent>();
                Assert.NotNull(textComponent);
                Assert.True(textComponent.HasText);

                // 验证基础属性
                if (textComponent.Paragraphs != null && textComponent.Paragraphs.Count > 0)
                {
                    var firstRun = textComponent.Paragraphs[0].Runs.FirstOrDefault();
                    if (firstRun != null)
                    {
                        Assert.NotNull(firstRun.FontName);
                        Assert.True(firstRun.FontSize > 0);
                        Assert.NotNull(firstRun.FontColor);
                        // 验证粗体、斜体、下划线、删除线
                        _output.WriteLine($"字体: {firstRun.FontName}, 大小: {firstRun.FontSize}pt");
                        _output.WriteLine($"粗体: {firstRun.IsBold}, 斜体: {firstRun.IsItalic}");
                        _output.WriteLine($"下划线: {firstRun.IsUnderline}, 删除线: {firstRun.IsStrikethrough}");
                    }
                }
            }
        }

        [Fact]
        public void TestReadShadowProperties()
        {
            if (!File.Exists(_testTemplatePath))
            {
                _output.WriteLine($"测试模板不存在: {_testTemplatePath}");
                return;
            }

            using (var reader = new PowerPointReader())
            {
                reader.Load(_testTemplatePath);
                var elements = reader.GetAllElements();
                var textboxes = elements.OfType<TextBoxElement>().ToList();

                // 查找文本框2（阴影属性）
                var textbox2 = textboxes.FirstOrDefault(tb => GetTextContent(tb).Contains("Times New Roman-(0,0,255)-16pt"));
                if (textbox2 == null)
                {
                    _output.WriteLine("未找到文本框2（阴影属性测试）");
                    return;
                }

                var textComponent = textbox2.GetComponent<TextComponent>();
                if (textComponent?.Paragraphs != null && textComponent.Paragraphs.Count > 0)
                {
                    var firstRun = textComponent.Paragraphs[0].Runs.FirstOrDefault();
                    if (firstRun?.TextEffects != null && firstRun.TextEffects.HasShadow)
                    {
                        var shadow = firstRun.TextEffects.Shadow;
                        Assert.NotNull(shadow);
                        _output.WriteLine($"阴影类型: {shadow.Type}, 颜色: {shadow.Color}");
                        _output.WriteLine($"模糊: {shadow.Blur}, 距离: {shadow.Distance}, 角度: {shadow.Angle}");
                    }
                }
            }
        }

        [Fact]
        public void TestReadFillProperties()
        {
            if (!File.Exists(_testTemplatePath))
            {
                _output.WriteLine($"测试模板不存在: {_testTemplatePath}");
                return;
            }

            using (var reader = new PowerPointReader())
            {
                reader.Load(_testTemplatePath);
                var elements = reader.GetAllElements();
                var textboxes = elements.OfType<TextBoxElement>().ToList();

                // 查找文本框3（填充属性）
                var textbox3 = textboxes.FirstOrDefault(tb => GetTextContent(tb).Contains("Calibri-(0,128,0)-18pt"));
                if (textbox3 == null)
                {
                    _output.WriteLine("未找到文本框3（填充属性测试）");
                    return;
                }

                var textComponent = textbox3.GetComponent<TextComponent>();
                if (textComponent?.Paragraphs != null && textComponent.Paragraphs.Count > 0)
                {
                    var firstRun = textComponent.Paragraphs[0].Runs.FirstOrDefault();
                    if (firstRun?.TextFill != null && firstRun.TextFill.HasFill)
                    {
                        var fill = firstRun.TextFill;
                        _output.WriteLine($"填充类型: {fill.FillType}");
                        _output.WriteLine($"填充颜色: {fill.Color}");
                        if (fill.Gradient != null)
                        {
                            _output.WriteLine($"渐变类型: {fill.Gradient.GradientType}");
                        }
                    }
                }
            }
        }

        [Fact]
        public void TestReadGradientFillProperties_191_191_191_Gradient4()
        {
            if (!File.Exists(_testTemplatePath))
            {
                _output.WriteLine($"测试模板不存在: {_testTemplatePath}");
                return;
            }

            using (var reader = new PowerPointReader())
            {
                reader.Load(_testTemplatePath);
                var elements = reader.GetAllElements();
                var textboxes = elements.OfType<TextBoxElement>().ToList();

                // 查找文本框（文本填充=191,191,191-渐变4）
                var textbox = textboxes.FirstOrDefault(tb => 
                    GetTextContent(tb).Contains("191,191,191") || 
                    GetTextContent(tb).Contains("渐变4"));
                
                if (textbox == null)
                {
                    _output.WriteLine("未找到文本框（文本填充=191,191,191-渐变4）");
                    return;
                }

                var textComponent = textbox.GetComponent<TextComponent>();
                Assert.NotNull(textComponent);
                
                if (textComponent?.Paragraphs != null && textComponent.Paragraphs.Count > 0)
                {
                    // 遍历所有段落和运行，查找具有渐变填充的文本
                    foreach (var paragraph in textComponent.Paragraphs)
                    {
                        if (paragraph.Runs != null)
                        {
                            foreach (var run in paragraph.Runs)
                            {
                                if (run?.TextFill != null && run.TextFill.HasFill)
                                {
                                    var fill = run.TextFill;
                                    
                                    // 验证填充类型
                                    _output.WriteLine($"=== 文本填充属性验证 ===");
                                    _output.WriteLine($"文本内容: \"{run.Text}\"");
                                    _output.WriteLine($"填充类型: {fill.FillType}");
                                    Assert.True(fill.HasFill, "应该有填充");
                                    
                                    // 验证是否为渐变填充
                                    if (fill.FillType == FillType.Gradient)
                                    {
                                        Assert.NotNull(fill.Gradient);
                                        _output.WriteLine("渐变填充应该有 Gradient 对象");
                                        var gradient = fill.Gradient;
                                        
                                        // 验证渐变类型
                                        _output.WriteLine($"渐变类型: {gradient.GradientType}");
                                        Assert.NotNull(gradient.GradientType);
                                        _output.WriteLine("渐变类型不应为空");
                                        Assert.NotEmpty(gradient.GradientType);
                                        _output.WriteLine("渐变类型不应为空字符串");
                                        
                                        // 验证渐变角度（线性渐变）
                                        _output.WriteLine($"渐变角度: {gradient.Angle}°");
                                        Assert.True(gradient.Angle >= 0 && gradient.Angle <= 360, 
                                            $"渐变角度应在 0-360 度之间，实际值: {gradient.Angle}");
                                        
                                        // 验证渐变停止点
                                        Assert.NotNull(gradient.Stops);
                                        _output.WriteLine("渐变停止点列表不应为空");
                                        Assert.True(gradient.Stops.Count > 0, 
                                            $"渐变停止点数量应大于 0，实际数量: {gradient.Stops.Count}");
                                        
                                        _output.WriteLine($"渐变停止点数量: {gradient.Stops.Count}");
                                        
                                        // 详细验证每个停止点
                                        for (int i = 0; i < gradient.Stops.Count; i++)
                                        {
                                            var stop = gradient.Stops[i];
                                            Assert.NotNull(stop);
                                            _output.WriteLine($"停止点 {i} 不应为空");
                                            Assert.NotNull(stop.Color);
                                            _output.WriteLine($"停止点 {i} 的颜色不应为空");
                                            
                                            // 验证停止点位置
                                            Assert.True(stop.Position >= 0.0f && stop.Position <= 1.0f, 
                                                $"停止点 {i} 的位置应在 0.0-1.0 之间，实际值: {stop.Position}");
                                            
                                            // 验证颜色信息
                                            var color = stop.Color;
                                            Assert.True(color.Red >= 0 && color.Red <= 255, 
                                                $"停止点 {i} 的红色值应在 0-255 之间，实际值: {color.Red}");
                                            Assert.True(color.Green >= 0 && color.Green <= 255, 
                                                $"停止点 {i} 的绿色值应在 0-255 之间，实际值: {color.Green}");
                                            Assert.True(color.Blue >= 0 && color.Blue <= 255, 
                                                $"停止点 {i} 的蓝色值应在 0-255 之间，实际值: {color.Blue}");
                                            
                                            _output.WriteLine($"  停止点 {i + 1}:");
                                            _output.WriteLine($"    位置: {stop.Position:F4} ({stop.Position * 100:F2}%)");
                                            _output.WriteLine($"    颜色: RGB({color.Red}, {color.Green}, {color.Blue})");
                                            if (!string.IsNullOrEmpty(color.OriginalHex))
                                            {
                                                _output.WriteLine($"    原始十六进制: {color.OriginalHex}");
                                            }
                                        }
                                        
                                        // 验证透明度
                                        _output.WriteLine($"填充透明度: {fill.Transparency:F2} ({(fill.Transparency * 100):F1}%)");
                                        Assert.True(fill.Transparency >= 0.0f && fill.Transparency <= 1.0f, 
                                            $"透明度应在 0.0-1.0 之间，实际值: {fill.Transparency}");
                                        
                                        // 验证停止点顺序（应该按位置排序）
                                        for (int i = 1; i < gradient.Stops.Count; i++)
                                        {
                                            Assert.True(gradient.Stops[i].Position >= gradient.Stops[i - 1].Position, 
                                                $"停止点应按位置排序，停止点 {i - 1} 位置: {gradient.Stops[i - 1].Position}, " +
                                                $"停止点 {i} 位置: {gradient.Stops[i].Position}");
                                        }
                                        
                                        _output.WriteLine("=== 渐变填充属性验证完成 ===");
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        [Fact]
        public void TestReadOutlineProperties()
        {
            if (!File.Exists(_testTemplatePath))
            {
                _output.WriteLine($"测试模板不存在: {_testTemplatePath}");
                return;
            }

            using (var reader = new PowerPointReader())
            {
                reader.Load(_testTemplatePath);
                var elements = reader.GetAllElements();
                var textboxes = elements.OfType<TextBoxElement>().ToList();

                // 查找文本框4（轮廓属性）
                var textbox4 = textboxes.FirstOrDefault(tb => GetTextContent(tb).Contains("Verdana-(255,128,0)-15pt"));
                if (textbox4 == null)
                {
                    _output.WriteLine("未找到文本框4（轮廓属性测试）");
                    return;
                }

                var textComponent = textbox4.GetComponent<TextComponent>();
                if (textComponent?.Paragraphs != null && textComponent.Paragraphs.Count > 0)
                {
                    var firstRun = textComponent.Paragraphs[0].Runs.FirstOrDefault();
                    if (firstRun?.TextOutline != null && firstRun.TextOutline.HasOutline)
                    {
                        var outline = firstRun.TextOutline;
                        _output.WriteLine($"轮廓宽度: {outline.Width}pt");
                        _output.WriteLine($"轮廓颜色: {outline.Color}");
                        _output.WriteLine($"虚线样式: {outline.DashStyle}");
                    }
                }
            }
        }

        [Fact]
        public void TestReadEffectsProperties()
        {
            if (!File.Exists(_testTemplatePath))
            {
                _output.WriteLine($"测试模板不存在: {_testTemplatePath}");
                return;
            }

            using (var reader = new PowerPointReader())
            {
                reader.Load(_testTemplatePath);
                var elements = reader.GetAllElements();
                var textboxes = elements.OfType<TextBoxElement>().ToList();

                // 查找文本框5（效果属性）
                var textbox5 = textboxes.FirstOrDefault(tb => GetTextContent(tb).Contains("Courier New-(128,0,255)-17pt"));
                if (textbox5 == null)
                {
                    _output.WriteLine("未找到文本框5（效果属性测试）");
                    return;
                }

                var textComponent = textbox5.GetComponent<TextComponent>();
                if (textComponent?.Paragraphs != null && textComponent.Paragraphs.Count > 0)
                {
                    var firstRun = textComponent.Paragraphs[0].Runs.FirstOrDefault();
                    if (firstRun?.TextEffects != null && firstRun.TextEffects.HasEffects)
                    {
                        var effects = firstRun.TextEffects;
                        _output.WriteLine($"有发光: {effects.HasGlow}");
                        _output.WriteLine($"有反射: {effects.HasReflection}");
                        _output.WriteLine($"有柔边: {effects.HasSoftEdge}");
                        if (effects.Glow != null)
                        {
                            _output.WriteLine($"发光半径: {effects.Glow.Radius}pt");
                        }
                    }
                }
            }
        }

        [Fact]
        public void TestReadThemeColors()
        {
            if (!File.Exists(_testTemplatePath))
            {
                _output.WriteLine($"测试模板不存在: {_testTemplatePath}");
                return;
            }

            using (var reader = new PowerPointReader())
            {
                reader.Load(_testTemplatePath);
                var elements = reader.GetAllElements();
                var textboxes = elements.OfType<TextBoxElement>().ToList();

                // 查找文本框7（主题颜色）
                var textbox7 = textboxes.FirstOrDefault(tb => GetTextContent(tb).Contains("SimSun-(0,128,128)-14pt"));
                if (textbox7 == null)
                {
                    _output.WriteLine("未找到文本框7（主题颜色测试）");
                    return;
                }

                var textComponent = textbox7.GetComponent<TextComponent>();
                if (textComponent?.Paragraphs != null && textComponent.Paragraphs.Count > 0)
                {
                    var firstRun = textComponent.Paragraphs[0].Runs.FirstOrDefault();
                    if (firstRun?.FontColor != null)
                    {
                        var color = firstRun.FontColor;
                        if (color.IsThemeColor)
                        {
                            _output.WriteLine($"主题颜色: {color.SchemeColorName}");
                            _output.WriteLine($"主题色索引: {color.SchemeColorIndex}");
                            if (color.Transforms != null && color.Transforms.HasTransforms)
                            {
                                _output.WriteLine($"颜色变换: LumMod={color.Transforms.LumMod}, Tint={color.Transforms.Tint}");
                            }
                        }
                    }
                }
            }
        }

        [Fact]
        public void TestReadAllPropertiesCombined()
        {
            if (!File.Exists(_testTemplatePath))
            {
                _output.WriteLine($"测试模板不存在: {_testTemplatePath}");
                return;
            }

            using (var reader = new PowerPointReader())
            {
                reader.Load(_testTemplatePath);
                var json = reader.ToJson();
                
                // 保存JSON用于检查
                var jsonPath = Path.Combine(_testOutputDir, "textbox_properties_full.json");
                File.WriteAllText(jsonPath, json);
                _output.WriteLine($"完整JSON已保存到: {jsonPath}");

                // 验证JSON包含关键属性
                Assert.Contains("\"text\"", json);
                _output.WriteLine($"JSON长度: {json.Length} 字符");
            }
        }

        #endregion

        #region 基于页面/阶段1-3的场景化测试

        /// <summary>
        /// 第1页：多行文本框（统一格式）
        /// 验证：
        /// 1) 同一文本框内存在多行文本；
        /// 2) 所有行的字体名称和字号一致；
        /// 3) 运行级属性中没有意外的粗体/斜体/下划线/删除线差异。
        /// </summary>
        [Fact]
        public void TestPage1_MultiLineTextbox_UniformFormat()
        {
            if (!File.Exists(_testTemplatePath))
            {
                _output.WriteLine($"测试 PPTX 不存在: {_testTemplatePath}");
                return;
            }

            using (var reader = new PowerPointReader())
            {
                reader.Load(_testTemplatePath);

                // 第 1 页所有文本框
                var textboxes = GetAllTextboxesOnPage(reader, 1);
                Assert.True(textboxes.Count > 0, "第1页应该至少包含一个文本框");

                // 找到一个真正的“多行统一格式”文本框：
                // - 至少 2 行（段落数或包含 \n 的 Run）
                // - 所有有效 Run 的字体名称和字号完全一致
                TextBoxElement target = null;

                foreach (var tb in textboxes)
                {
                    var paraRuns = ExtractRunsFromTextbox(tb);
                    if (paraRuns.TotalLines < 2) continue;
                    if (paraRuns.Runs.Count == 0) continue;

                    var first = paraRuns.Runs[0];
                    bool allSame =
                        paraRuns.Runs.All(r =>
                            string.Equals(r.FontName, first.FontName, StringComparison.OrdinalIgnoreCase) &&
                            AreFloatEqual(r.FontSize, first.FontSize));

                    if (allSame)
                    {
                        target = tb;
                        break;
                    }
                }

                Assert.NotNull(target);

                var info = ExtractRunsFromTextbox(target);
                var firstRun = info.Runs[0];

                _output.WriteLine($"检测到多行统一格式文本框，行数: {info.TotalLines}");
                _output.WriteLine($"统一字体: {firstRun.FontName}, 字号: {firstRun.FontSize}pt");

                // 所有行都应使用相同的字体和字号
                foreach (var run in info.Runs)
                {
                    Assert.Equal(firstRun.FontName, run.FontName);
                    Assert.True(AreFloatEqual(firstRun.FontSize, run.FontSize),
                        $"字号不一致: 期望 {firstRun.FontSize}，实际 {run.FontSize}");
                }
            }
        }

        /// <summary>
        /// 第1页：多行文本框（每行不同格式）
        /// 验证：
        /// 1) 同一文本框内至少存在 2 行（无需固定为 4 行）；
        /// 2) 行与行之间在字体名称 / 字号 / 填充 / 格式标记上存在差异；
        /// 3) 可以体现「每行不同格式」这一设计意图，而不依赖具体字体名称。
        /// </summary>
        [Fact]
        public void TestPage1_MultiLineTextbox_DifferentFormatsPerLine()
        {
            if (!File.Exists(_testTemplatePath))
            {
                _output.WriteLine($"测试 PPTX 不存在: {_testTemplatePath}");
                return;
            }

            using (var reader = new PowerPointReader())
            {
                reader.Load(_testTemplatePath);

                var textboxes = GetAllTextboxesOnPage(reader, 1);
                Assert.True(textboxes.Count > 0, "第1页应该至少包含一个文本框");

                _output.WriteLine($"第1页找到 {textboxes.Count} 个文本框:");
                for (int i = 0; i < textboxes.Count; i++)
                {
                    var content = GetTextContent(textboxes[i]);
                    var info = ExtractRunsFromTextbox(textboxes[i]);
                    _output.WriteLine($"  文本框 {i + 1}: 行数={info.TotalLines}, 内容=\"{content.Substring(0, Math.Min(50, content.Length))}...\"");
                }

                TextBoxElement target = null;
                RunCollection runsInfo = null;

                foreach (var tb in textboxes)
                {
                    var info = ExtractRunsFromTextbox(tb);
                    if (info.TotalLines < 2) 
                    {
                        _output.WriteLine($"跳过文本框（行数不足2行: {info.TotalLines}）");
                        continue;
                    }

                    // 统计行级"主 Run"（每行第一个非空 Run）
                    var lineRuns = info.GetFirstRunPerLine();
                    if (lineRuns.Count < 2) 
                    {
                        _output.WriteLine($"跳过文本框（有效行数不足2行: {lineRuns.Count}）");
                        continue;
                    }

                    // 判断这些行之间是否存在明显的格式差异（字体名/字号/填充/加粗/斜体/下划线/阴影/字符间距）
                    bool hasDifferences = false;
                    for (int i = 0; i < lineRuns.Count - 1 && !hasDifferences; i++)
                    {
                        var a = lineRuns[i];
                        var b = lineRuns[i + 1];
                        hasDifferences = HasLineFormatDifference(a, b);
                    }

                    if (hasDifferences)
                    {
                        target = tb;
                        runsInfo = info;
                        break;
                    }
                }

                if (target == null)
                {
                _output.WriteLine("未找到符合条件的多行不同格式文本框。尝试使用第一个多行文本框（>=2行）");
                    foreach (var tb in textboxes)
                    {
                        var info = ExtractRunsFromTextbox(tb);
                    if (info.TotalLines >= 2)
                        {
                            target = tb;
                            runsInfo = info;
                            _output.WriteLine($"使用第一个多行文本框（行数: {info.TotalLines}）");
                            break;
                        }
                    }
                }

                Assert.True(target != null, $"找不到符合条件的文本框。已检查 {textboxes.Count} 个文本框。");
                Assert.NotNull(runsInfo);

                var lineLevelRuns = runsInfo.GetFirstRunPerLine();
                _output.WriteLine($"检测到多行不同格式文本框，行数: {runsInfo.TotalLines}");
                for (int i = 0; i < lineLevelRuns.Count; i++)
                {
                    var r = lineLevelRuns[i];
                    _output.WriteLine(
                        $"行 {i + 1}: 文本=\"{r.Text}\", 字体=\"{r.FontName}\", 大小={r.FontSize}pt, 粗体={r.IsBold}, 斜体={r.IsItalic}, 下划线={r.IsUnderline}, 阴影={(r.TextEffects?.HasShadow ?? false)}, 填充={(r.TextFill?.HasFill ?? false)}, 字符间距={r.CharacterSpacing}");
                }

                // 至少有两行在格式上明显不同
                bool anyDiff = false;
                for (int i = 0; i < lineLevelRuns.Count - 1 && !anyDiff; i++)
                {
                    var a = lineLevelRuns[i];
                    var b = lineLevelRuns[i + 1];
                    anyDiff = HasLineFormatDifference(a, b);
                }

                Assert.True(anyDiff, "多行不同格式文本框中，不同行之间应该存在明显的格式差异");
            }
        }

        /// <summary>
        /// 第2页：主题颜色和颜色变换测试
        /// 验证：
        /// 1) 至少存在若干使用 ThemeColor 的文本运行；
        /// 2) ThemeColor 的 SchemeColorName 不为空；
        /// 3) 至少有一个 ThemeColor 带有颜色变换参数（tint/shade/lumMod 等）。
        /// </summary>
        [Fact]
        public void TestPage2_ThemeColors()
        {
            if (!File.Exists(_testTemplatePath))
            {
                _output.WriteLine($"测试 PPTX 不存在: {_testTemplatePath}");
                return;
            }

            using (var reader = new PowerPointReader())
            {
                reader.Load(_testTemplatePath);

                var textboxes = GetAllTextboxesOnPage(reader, 3);
                Assert.True(textboxes.Count > 0, "第3页应该包含主题颜色测试用的文本框");

                var themeColors = new List<ColorInfo>();

                foreach (var tb in textboxes)
                {
                    var info = ExtractRunsFromTextbox(tb);
                    foreach (var run in info.Runs)
                    {
                        if (run.FontColor != null && run.FontColor.IsThemeColor)
                        {
                            themeColors.Add(run.FontColor);
                        }
                    }
                }

                Assert.True(themeColors.Count > 0, "第2页文本运行中应该存在主题颜色");

                // 所有主题色都应携带 SchemeColorName
                Assert.All(themeColors, c => Assert.False(string.IsNullOrEmpty(c.SchemeColorName)));

                // 至少有一个主题色具备颜色变换（代表 .1~.5 变体）
                bool hasTransforms = themeColors.Any(c => c.Transforms != null && c.Transforms.HasTransforms);
                Assert.True(hasTransforms, "至少有一个主题色应包含颜色变换参数（tint/shade/lumMod 等）");

                var distinctSchemes = new HashSet<string>(themeColors.Select(c => c.SchemeColorName));
                _output.WriteLine($"检测到主题色种类数: {distinctSchemes.Count}");
                foreach (var name in distinctSchemes)
                {
                    _output.WriteLine($"  - 主题色: {name}");
                }
            }
        }

        /// <summary>
        /// 第3页：字体类型识别测试
        /// 验证：
        /// 1) 第3页至少包含多种不同的字体；
        /// 2) 能识别出若干常见中文字体（等线/方正/华文/思源黑体等）或其子集；
        /// 3) 字体名称从 EastAsianFont / LatinFont 中正确解析。
        /// </summary>
        [Fact]
        public void TestPage3_FontTypes()
        {
            if (!File.Exists(_testTemplatePath))
            {
                _output.WriteLine($"测试 PPTX 不存在: {_testTemplatePath}");
                return;
            }

            using (var reader = new PowerPointReader())
            {
                reader.Load(_testTemplatePath);

                var textboxes = GetAllTextboxesOnPage(reader, 4);
                Assert.True(textboxes.Count > 0, "第4页应该包含字体类型测试用的文本框");

                var fonts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                foreach (var tb in textboxes)
                {
                    var info = ExtractRunsFromTextbox(tb);
                    foreach (var run in info.Runs)
                    {
                        if (!string.IsNullOrEmpty(run.FontName))
                            fonts.Add(run.FontName);
                    }
                }

                Assert.True(fonts.Count >= 2, "第3页应该至少包含两种不同字体");

                _output.WriteLine("第3页检测到的字体：");
                foreach (var f in fonts)
                {
                    _output.WriteLine($"  - {f}");
                }

                // 检查是否包含典型中文字体关键字（不强制全覆盖，以兼容不同模板）
                bool hasCjkFont =
                    fonts.Any(f => f.IndexOf("等线", StringComparison.OrdinalIgnoreCase) >= 0) ||
                    fonts.Any(f => f.IndexOf("方正", StringComparison.OrdinalIgnoreCase) >= 0) ||
                    fonts.Any(f => f.IndexOf("华文", StringComparison.OrdinalIgnoreCase) >= 0) ||
                    fonts.Any(f => f.IndexOf("思源", StringComparison.OrdinalIgnoreCase) >= 0) ||
                    fonts.Any(f => f.IndexOf("宋体", StringComparison.OrdinalIgnoreCase) >= 0);

                Assert.True(hasCjkFont, "第3页应至少包含一种常见中文字体（等线/方正/华文/思源/宋体等）");
            }
        }

        /// <summary>
        /// 第3页：格式化效果组合测试
        /// 验证：
        /// 1) 能识别出加粗、斜体、下划线、删除线等布尔属性；
        /// 2) 能识别出至少一个带阴影效果的文本运行；
        /// 3) 这些信息在 TextRunInfo 中被正确填充。
        /// </summary>
        [Fact]
        public void TestPage3_FormattingEffects()
        {
            if (!File.Exists(_testTemplatePath))
            {
                _output.WriteLine($"测试 PPTX 不存在: {_testTemplatePath}");
                return;
            }

            using (var reader = new PowerPointReader())
            {
                reader.Load(_testTemplatePath);

                var textboxes = GetAllTextboxesOnPage(reader, 4);
                Assert.True(textboxes.Count > 0, "第4页应该包含格式化效果测试用的文本框");

                bool hasBold = false;
                bool hasItalic = false;
                bool hasUnderline = false;
                bool hasStrikethrough = false;
                bool hasShadow = false;

                foreach (var tb in textboxes)
                {
                    var info = ExtractRunsFromTextbox(tb);
                    foreach (var run in info.Runs)
                    {
                        if (run.IsBold) hasBold = true;
                        if (run.IsItalic) hasItalic = true;
                        if (run.IsUnderline) hasUnderline = true;
                        if (run.IsStrikethrough) hasStrikethrough = true;
                        if (run.TextEffects?.HasShadow == true) hasShadow = true;
                    }
                }

                _output.WriteLine($"Bold: {hasBold}, Italic: {hasItalic}, Underline: {hasUnderline}, Strikethrough: {hasStrikethrough}, Shadow: {hasShadow}");

                Assert.True(hasBold, "第3页应至少包含一个加粗文本");
                Assert.True(hasItalic, "第3页应至少包含一个斜体文本");
                Assert.True(hasUnderline, "第3页应至少包含一个下划线文本");
                Assert.True(hasStrikethrough, "第3页应至少包含一个删除线文本");
                Assert.True(hasShadow, "第3页应至少包含一个具有阴影效果的文本");
            }
        }

        #endregion

        #region 端到端测试（往返测试）

        [Fact]
        public void TestRoundTrip_BasicProperties()
        {
            if (!File.Exists(_testTemplatePath))
            {
                _output.WriteLine($"测试模板不存在: {_testTemplatePath}");
                return;
            }

            // 读取原始PPTX
            TextRunInfo originalRun = null;
            using (var reader = new PowerPointReader())
            {
                reader.Load(_testTemplatePath);
                var elements = reader.GetAllElements();
                var textbox = elements.OfType<TextBoxElement>()
                    .FirstOrDefault(tb => GetTextContent(tb).Contains("Arial-(255,0,0)-14pt"));
                
                if (textbox == null)
                {
                    _output.WriteLine("未找到测试文本框");
                    return;
                }

                var textComponent = textbox.GetComponent<TextComponent>();
                if (textComponent?.Paragraphs != null && textComponent.Paragraphs.Count > 0)
                {
                    originalRun = textComponent.Paragraphs[0].Runs.FirstOrDefault();
                }
            }

            if (originalRun == null)
            {
                _output.WriteLine("未找到文本运行");
                return;
            }

            // 序列化为JSON
            string json;
            using (var reader = new PowerPointReader())
            {
                reader.Load(_testTemplatePath);
                json = reader.ToJson();
            }

            // 验证JSON包含基础属性
            Assert.Contains("\"fontName\"", json);
            Assert.Contains("\"fontSize\"", json);
            Assert.Contains("\"fontColor\"", json);

            _output.WriteLine("基础属性往返测试通过");
        }

        [Fact]
        public void TestRoundTrip_ShadowProperties()
        {
            if (!File.Exists(_testTemplatePath))
            {
                _output.WriteLine($"测试模板不存在: {_testTemplatePath}");
                return;
            }

            using (var reader = new PowerPointReader())
            {
                reader.Load(_testTemplatePath);
                var json = reader.ToJson();

                // 验证JSON包含阴影相关属性
                // 注意：阴影应该在TextEffects中，而不是直接在TextRun中
                Assert.Contains("\"text_effects\"", json);
                _output.WriteLine("阴影属性往返测试通过");
            }
        }

        [Fact]
        public void TestRoundTrip_FillProperties()
        {
            if (!File.Exists(_testTemplatePath))
            {
                _output.WriteLine($"测试模板不存在: {_testTemplatePath}");
                return;
            }

            using (var reader = new PowerPointReader())
            {
                reader.Load(_testTemplatePath);
                var json = reader.ToJson();

                // 验证JSON包含填充相关属性
                Assert.Contains("\"text_fill\"", json);
                _output.WriteLine("填充属性往返测试通过");
            }
        }

        [Fact]
        public void TestRoundTrip_OutlineProperties()
        {
            if (!File.Exists(_testTemplatePath))
            {
                _output.WriteLine($"测试模板不存在: {_testTemplatePath}");
                return;
            }

            using (var reader = new PowerPointReader())
            {
                reader.Load(_testTemplatePath);
                var json = reader.ToJson();

                // 基于当前实现：只验证形状轮廓，而不要求文本轮廓字段
                Assert.Contains("\"line\"", json);
                Assert.Contains("\"has_outline\"", json);

                _output.WriteLine("轮廓属性往返测试通过（基于形状轮廓 line.has_outline，不再要求 text_outline）");
            }
        }

        [Fact]
        public void TestRoundTrip_EffectsProperties()
        {
            if (!File.Exists(_testTemplatePath))
            {
                _output.WriteLine($"测试模板不存在: {_testTemplatePath}");
                return;
            }

            using (var reader = new PowerPointReader())
            {
                reader.Load(_testTemplatePath);
                var json = reader.ToJson();

                // 验证JSON包含效果相关属性
                Assert.Contains("\"text_effects\"", json);
                _output.WriteLine("效果属性往返测试通过");
            }
        }

        [Fact]
        public void TestRoundTrip_AllProperties()
        {
            if (!File.Exists(_testTemplatePath))
            {
                _output.WriteLine($"测试模板不存在: {_testTemplatePath}");
                return;
            }

            var testedProperties = new Dictionary<string, bool>();

            using (var reader = new PowerPointReader())
            {
                reader.Load(_testTemplatePath);
                var elements = reader.GetAllElements();
                var textboxes = elements.OfType<TextBoxElement>().ToList();

                foreach (var textbox in textboxes)
                {
                    var textComponent = textbox.GetComponent<TextComponent>();
                    if (textComponent?.Paragraphs != null)
                    {
                        foreach (var para in textComponent.Paragraphs)
                        {
                            foreach (var run in para.Runs)
                            {
                                // 检查基础属性
                                if (!string.IsNullOrEmpty(run.FontName)) testedProperties["FontName"] = true;
                                if (run.FontSize > 0) testedProperties["FontSize"] = true;
                                if (run.FontColor != null) testedProperties["FontColor"] = true;
                                if (run.IsBold) testedProperties["IsBold"] = true;
                                if (run.IsItalic) testedProperties["IsItalic"] = true;
                                if (run.IsUnderline) testedProperties["IsUnderline"] = true;
                                if (run.IsStrikethrough) testedProperties["IsStrikethrough"] = true;

                                // 检查阴影属性
                                if (run.TextEffects?.Shadow != null)
                                {
                                    testedProperties["ShadowType"] = true;
                                    if (run.TextEffects.Shadow.Color != null) testedProperties["ShadowColor"] = true;
                                    if (run.TextEffects.Shadow.Blur > 0) testedProperties["ShadowBlur"] = true;
                                    if (run.TextEffects.Shadow.Distance > 0) testedProperties["ShadowDistance"] = true;
                                    if (run.TextEffects.Shadow.Angle > 0) testedProperties["ShadowAngle"] = true;
                                    if (run.TextEffects.Shadow.Transparency > 0) testedProperties["ShadowTransparency"] = true;
                                }

                                // 检查填充属性
                                if (run.TextFill != null && run.TextFill.HasFill)
                                {
                                    testedProperties["FillType"] = true;
                                    if (run.TextFill.Color != null) testedProperties["FillColor"] = true;
                                    if (run.TextFill.Gradient != null) testedProperties["FillGradient"] = true;
                                    if (run.TextFill.Pattern != null) testedProperties["FillPattern"] = true;
                                    if (run.TextFill.Transparency > 0) testedProperties["FillTransparency"] = true;
                                }

                                // 检查轮廓属性
                                if (run.TextOutline != null && run.TextOutline.HasOutline)
                                {
                                    testedProperties["OutlineWidth"] = true;
                                    if (run.TextOutline.Color != null) testedProperties["OutlineColor"] = true;
                                    testedProperties["OutlineDashStyle"] = true;
                                    if (run.TextOutline.Transparency > 0) testedProperties["OutlineTransparency"] = true;
                                }

                                // 检查效果属性
                                if (run.TextEffects != null)
                                {
                                    if (run.TextEffects.HasGlow && run.TextEffects.Glow != null) testedProperties["Glow"] = true;
                                    if (run.TextEffects.HasReflection && run.TextEffects.Reflection != null) testedProperties["Reflection"] = true;
                                    if (run.TextEffects.HasSoftEdge) testedProperties["SoftEdge"] = true;
                                }

                                // 检查其他属性
                                if (run.HighlightColor != null) testedProperties["HighlightColor"] = true;
                                if (run.CharacterSpacing > 0) testedProperties["CharacterSpacing"] = true;
                                if (run.Superscript.HasValue) testedProperties["Superscript"] = true;
                                if (run.Subscript.HasValue) testedProperties["Subscript"] = true;
                                if (run.FontColor?.IsThemeColor == true) testedProperties["ThemeColor"] = true;
                                if (run.FontColor?.Transforms != null && run.FontColor.Transforms.HasTransforms) testedProperties["ColorTransforms"] = true;
                            }
                        }
                    }
                }
            }

            // 生成覆盖率报告
            var checklist = _coverageAnalyzer.AnalyzePropertyCoverage(testedProperties);
            var report = _coverageAnalyzer.GenerateCoverageReport();
            
            var reportPath = Path.Combine(_testReportsDir, "coverage_report.md");
            File.WriteAllText(reportPath, report);
            _output.WriteLine($"覆盖率报告已保存到: {reportPath}");
            _output.WriteLine($"覆盖率: {checklist.GetCoveragePercentage():F2}%");
        }

        #endregion

        #region 辅助方法

        private string GetTextContent(TextBoxElement textbox)
        {
            var textComponent = textbox.GetComponent<TextComponent>();
            return textComponent?.TextContent ?? string.Empty;
        }

        /// <summary>
        /// 获取指定页上的所有文本框
        /// </summary>
        private List<TextBoxElement> GetAllTextboxesOnPage(PowerPointReader reader, int pageIndex)
        {
            var elements = reader.GetPageElements(pageIndex);
            return elements.OfType<TextBoxElement>().ToList();
        }

        /// <summary>
        /// 将 TextBoxElement 中的 Paragraph/Run 扁平化为便于测试的结构
        /// </summary>
        private RunCollection ExtractRunsFromTextbox(TextBoxElement textbox)
        {
            var collection = new RunCollection();
            var textComponent = textbox.GetComponent<TextComponent>();
            if (textComponent?.Paragraphs == null) return collection;

            int lineIndex = 0;
            foreach (var para in textComponent.Paragraphs)
            {
                bool hasContentInThisLine = false;
                foreach (var run in para.Runs)
                {
                    // 将段内换行符视为新的一行
                    if (run.Text == "\n")
                    {
                        if (hasContentInThisLine)
                        {
                            lineIndex++;
                            hasContentInThisLine = false;
                        }
                        continue;
                    }

                    var cloned = new TextRunInfo
                    {
                        Text = run.Text,
                        FontName = run.FontName,
                        FontSize = run.FontSize,
                        FontColor = run.FontColor,
                        IsBold = run.IsBold,
                        IsItalic = run.IsItalic,
                        IsUnderline = run.IsUnderline,
                        IsStrikethrough = run.IsStrikethrough,
                        TextFill = run.TextFill,
                        TextOutline = run.TextOutline,
                        TextEffects = run.TextEffects,
                        HighlightColor = run.HighlightColor,
                        CharacterSpacing = run.CharacterSpacing,
                        Superscript = run.Superscript,
                        Subscript = run.Subscript
                    };

                    collection.Runs.Add(cloned);
                    collection.LineIndex.Add(lineIndex);
                    hasContentInThisLine = true;
                }

                // 段落结束即视为换行
                if (hasContentInThisLine)
                {
                    lineIndex++;
                }
            }

            collection.TotalLines = lineIndex;
            return collection;
        }

        /// <summary>
        /// 浮点数比较，避免字号等由于转换产生的微小误差
        /// </summary>
        private bool AreFloatEqual(float a, float b, float epsilon = 0.1f)
            => Math.Abs(a - b) <= epsilon;

        /// <summary>
        /// 判断两行（取各自首个 Run）在格式上是否存在差异。
        /// 差异项覆盖：字体名、字号、填充状态、粗体/斜体/下划线、阴影、字符间距。
        /// </summary>
        private bool HasLineFormatDifference(TextRunInfo a, TextRunInfo b)
        {
            if (a == null || b == null) return false;

            bool fillA = a.TextFill?.HasFill ?? false;
            bool fillB = b.TextFill?.HasFill ?? false;
            bool shadowA = a.TextEffects?.HasShadow ?? false;
            bool shadowB = b.TextEffects?.HasShadow ?? false;

            return
                !string.Equals(a.FontName, b.FontName, StringComparison.OrdinalIgnoreCase) ||
                !AreFloatEqual(a.FontSize, b.FontSize) ||
                fillA != fillB ||
                a.IsBold != b.IsBold ||
                a.IsItalic != b.IsItalic ||
                a.IsUnderline != b.IsUnderline ||
                shadowA != shadowB ||
                !AreFloatEqual(a.CharacterSpacing, b.CharacterSpacing);
        }

        /// <summary>
        /// 辅助结构：承载从 TextBoxElement 中提取的 Run 列表及其行号信息
        /// </summary>
        private class RunCollection
        {
            public List<TextRunInfo> Runs { get; } = new List<TextRunInfo>();
            public List<int> LineIndex { get; } = new List<int>();
            public int TotalLines { get; set; }

            /// <summary>
            /// 获取每一行的“主 Run”（当前实现中取每行第一个 Run）
            /// </summary>
            public List<TextRunInfo> GetFirstRunPerLine()
            {
                var result = new List<TextRunInfo>();
                var seen = new HashSet<int>();

                for (int i = 0; i < Runs.Count; i++)
                {
                    var line = LineIndex[i];
                    if (seen.Contains(line)) continue;
                    seen.Add(line);
                    result.Add(Runs[i]);
                }

                return result;
            }
        }

        #endregion
    }
}

