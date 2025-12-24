using System;
using System.IO;
using Xunit;
using OfficeHelperOpenXml.Api;

namespace OfficeHelperOpenXml.Test
{
    /// <summary>
    /// 模板文件结构分析测试
    /// 阶段一：分析 26xdemo2.pptx 模板文件结构
    /// </summary>
    public class TemplateAnalysisTests
    {
        private static string TemplatePath => TestPaths.Template26xdemo2Path;

        /// <summary>
        /// 测试模板文件是否存在
        /// </summary>
        [Fact]
        public void TestTemplateFileExists()
        {
            Assert.True(File.Exists(TemplatePath), $"模板文件不存在: {TemplatePath}");
        }

        /// <summary>
        /// 分析模板文件结构
        /// </summary>
        [Fact]
        public void TestAnalyzeTemplateStructure()
        {
            if (!File.Exists(TemplatePath))
            {
                Assert.True(false, $"模板文件不存在: {TemplatePath}");
                return;
            }

            using (var reader = PowerPointReaderFactory.CreateReader(TemplatePath, out bool success))
            {
                Assert.True(success, "无法创建 PowerPointReader");
                Assert.NotNull(reader);

                var info = reader.PresentationInfo;
                Assert.NotNull(info);
                Assert.True(string.IsNullOrEmpty(info.Error), $"读取模板文件失败: {info.Error}");

                // 验证基本信息
                Assert.True(info.SlideWidth > 0, "幻灯片宽度应该大于 0");
                Assert.True(info.SlideHeight > 0, "幻灯片高度应该大于 0");
                Assert.True(info.SlideCount >= 0, "幻灯片数量应该 >= 0");

                // 输出分析报告
                Console.WriteLine("========== 模板文件结构分析 ==========");
                Console.WriteLine($"文件路径: {TemplatePath}");
                Console.WriteLine($"幻灯片尺寸: {info.SlideWidth} x {info.SlideHeight} cm");
                Console.WriteLine($"幻灯片数量: {info.SlideCount}");
                Console.WriteLine($"母版幻灯片数量: {info.MasterSlides?.Count ?? 0}");
                Console.WriteLine($"内容幻灯片数量: {info.Slides?.Count ?? 0}");
                Console.WriteLine($"默认文本样式: {(info.DefaultTextStyle != null ? "已提取" : "未提取")}");
                Console.WriteLine($"母版样式数量: {info.SlideMasterStyles?.Count ?? 0}");

                // 分析每张幻灯片
                if (info.Slides != null && info.Slides.Count > 0)
                {
                    Console.WriteLine("\n========== 内容幻灯片详情 ==========");
                    for (int i = 0; i < info.Slides.Count; i++)
                    {
                        var slide = info.Slides[i];
                        Console.WriteLine($"幻灯片 {slide.SlideNumber}:");
                        Console.WriteLine($"  - 索引: {slide.SlideIndex}");
                        Console.WriteLine($"  - 元素数量: {slide.Elements?.Count ?? 0}");
                    }
                }

                // 分析母版幻灯片
                if (info.MasterSlides != null && info.MasterSlides.Count > 0)
                {
                    Console.WriteLine("\n========== 母版幻灯片详情 ==========");
                    for (int i = 0; i < info.MasterSlides.Count; i++)
                    {
                        var master = info.MasterSlides[i];
                        Console.WriteLine($"母版 {master.PageNumber}:");
                        Console.WriteLine($"  - 元素数量: {master.Shapes?.Count ?? 0}");
                    }
                }

                // 分析母版样式
                if (info.SlideMasterStyles != null && info.SlideMasterStyles.Count > 0)
                {
                    Console.WriteLine("\n========== 母版样式详情 ==========");
                    for (int i = 0; i < info.SlideMasterStyles.Count; i++)
                    {
                        var style = info.SlideMasterStyles[i];
                        Console.WriteLine($"母版样式 {i + 1}:");
                        Console.WriteLine($"  - ID: {style.MasterId}");
                        Console.WriteLine($"  - 背景填充: {(style.Background != null ? "有" : "无")}");
                        Console.WriteLine($"  - 颜色方案: {(style.ColorScheme != null ? "有" : "无")}");
                    }
                }

                Console.WriteLine("\n========== 分析完成 ==========");
            }
        }

        /// <summary>
        /// 识别母版幻灯片和内容幻灯片
        /// </summary>
        [Fact]
        public void TestIdentifyMasterAndContentSlides()
        {
            if (!File.Exists(TemplatePath))
            {
                Assert.True(false, $"模板文件不存在: {TemplatePath}");
                return;
            }

            using (var reader = PowerPointReaderFactory.CreateReader(TemplatePath, out bool success))
            {
                Assert.True(success);
                var info = reader.PresentationInfo;
                Assert.NotNull(info);

                // 验证母版幻灯片存在
                Assert.NotNull(info.MasterSlides);
                Assert.True(info.MasterSlides.Count > 0, "模板应该至少有一个母版幻灯片");

                // 验证内容幻灯片
                Assert.NotNull(info.Slides);
                Console.WriteLine($"母版幻灯片数量: {info.MasterSlides.Count}");
                Console.WriteLine($"内容幻灯片数量: {info.Slides.Count}");
            }
        }

        /// <summary>
        /// 记录模板的样式信息
        /// </summary>
        [Fact]
        public void TestExtractTemplateStyles()
        {
            if (!File.Exists(TemplatePath))
            {
                Assert.True(false, $"模板文件不存在: {TemplatePath}");
                return;
            }

            using (var reader = PowerPointReaderFactory.CreateReader(TemplatePath, out bool success))
            {
                Assert.True(success);
                var info = reader.PresentationInfo;
                Assert.NotNull(info);

                // 验证默认文本样式
                Assert.NotNull(info.DefaultTextStyle);
                Console.WriteLine("默认文本样式:");
                var level1 = info.DefaultTextStyle.Levels?.Level1;
                Console.WriteLine($"  - 字体: {(level1 != null && !string.IsNullOrEmpty(level1.FontEa) ? level1.FontEa : "未设置")}");
                Console.WriteLine($"  - 字号: {(level1 != null ? level1.FontSize : 0)}");
                Console.WriteLine($"  - 颜色: {(level1 != null && !string.IsNullOrEmpty(level1.FontColor) ? level1.FontColor : "未设置")}");

                // 验证母版样式
                Assert.NotNull(info.SlideMasterStyles);
                Console.WriteLine($"\n母版样式数量: {info.SlideMasterStyles.Count}");

                // 生成样式报告
                var json = reader.ToJson();
                Assert.NotNull(json);
                Assert.True(json.Length > 0, "JSON 输出不应该为空");

                // 可选：保存 JSON 到文件用于分析
                var jsonPath = "template_analysis_output.json";
                try
                {
                    File.WriteAllText(jsonPath, json);
                    Console.WriteLine($"\n模板分析 JSON 已保存到: {jsonPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"保存 JSON 失败: {ex.Message}");
                }
            }
        }
    }
}

