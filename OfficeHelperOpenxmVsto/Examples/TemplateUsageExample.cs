using System;
using System.Collections.Generic;
using System.IO;
using OfficeHelperOpenXml.Api;
using OfficeHelperOpenXml.Api.PowerPoint;
using OfficeHelperOpenXml.Models.Json;

namespace OfficeHelperOpenXml.Examples
{
    /// <summary>
    /// 模板使用示例
    /// 演示如何使用 26xdemo2.pptx 模板创建新的 PPTX 文件
    /// </summary>
    public class TemplateUsageExample
    {
        /// <summary>
        /// 示例 1: 使用便捷方法从 JSON 文件生成 PPTX
        /// </summary>
        public static void Example1_SimpleUsage()
        {
            string templatePath = "26xdemo2.pptx";
            string jsonDataPath = "content.json";
            string outputPath = "output.pptx";

            // 读取 JSON 数据
            if (!File.Exists(jsonDataPath))
            {
                Console.WriteLine($"JSON 文件不存在: {jsonDataPath}");
                return;
            }

            string jsonData = File.ReadAllText(jsonDataPath);

            // 使用便捷方法生成 PPTX
            bool success = OfficeHelperWrapper.WritePowerPointFromJson(
                templatePath, jsonData, outputPath);

            if (success)
            {
                Console.WriteLine($"成功生成文件: {outputPath}");
            }
            else
            {
                Console.WriteLine("生成文件失败");
            }
        }

        /// <summary>
        /// 示例 2: 使用 PowerPointWriter API 创建简单幻灯片
        /// </summary>
        public static void Example2_CreateSimpleSlide()
        {
            string templatePath = "26xdemo2.pptx";
            string outputPath = "simple_output.pptx";

            using (var writer = PowerPointWriterFactory.CreateWriter())
            {
                // 1. 打开模板
                if (!writer.OpenFromTemplate(templatePath))
                {
                    Console.WriteLine("打开模板失败");
                    return;
                }

                // 2. 清除内容幻灯片
                writer.ClearAllContentSlides();

                // 3. 创建简单的 JSON 数据
                var jsonData = new PresentationJsonData
                {
                    ContentSlides = new List<SlideJsonData>
                    {
                        new SlideJsonData
                        {
                            PageNumber = 1,
                            Title = "简单示例",
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
                                            Content = "这是一个简单的示例",
                                            Font = "Arial",
                                            FontSize = 24,
                                            FontColor = "RGB(0,0,0)",
                                            FontBold = 1
                                        }
                                    },
                                    Fill = new FillJsonData
                                    {
                                        Color = "RGB(255,255,255)",
                                        Opacity = 1.0f
                                    }
                                }
                            }
                        }
                    }
                };

                // 4. 写入数据
                if (!writer.WriteFromJsonData(jsonData))
                {
                    Console.WriteLine("写入数据失败");
                    return;
                }

                // 5. 保存文件
                if (!writer.SaveAs(outputPath))
                {
                    Console.WriteLine("保存文件失败");
                    return;
                }

                Console.WriteLine($"成功生成文件: {outputPath}");
            }
        }

        /// <summary>
        /// 示例 3: 创建包含多种形状的复杂幻灯片
        /// </summary>
        public static void Example3_CreateComplexSlide()
        {
            string templatePath = "26xdemo2.pptx";
            string outputPath = "complex_output.pptx";

            using (var writer = PowerPointWriterFactory.CreateWriter())
            {
                writer.OpenFromTemplate(templatePath);
                writer.ClearAllContentSlides();

                var jsonData = new PresentationJsonData
                {
                    ContentSlides = new List<SlideJsonData>
                    {
                        new SlideJsonData
                        {
                            PageNumber = 1,
                            Title = "复杂示例",
                            Shapes = new List<ShapeJsonData>
                            {
                                // 标题文本框
                                new ShapeJsonData
                                {
                                    Type = "textbox",
                                    Name = "Title",
                                    Box = "1,1,22,2",
                                    HasText = 1,
                                    Text = new List<TextRunJsonData>
                                    {
                                        new TextRunJsonData
                                        {
                                            Content = "复杂幻灯片示例",
                                            Font = "Arial",
                                            FontSize = 32,
                                            FontColor = "RGB(0,0,128)",
                                            FontBold = 1
                                        }
                                    }
                                },
                                // 矩形
                                new ShapeJsonData
                                {
                                    Type = "autoshape",
                                    Name = "Rectangle",
                                    Box = "2,4,8,6",
                                    SpecialType = "rectangle",
                                    Fill = new FillJsonData
                                    {
                                        Color = "RGB(200,200,200)",
                                        Opacity = 1.0f
                                    },
                                    Line = new LineJsonData
                                    {
                                        HasOutline = 1,
                                        Color = "RGB(0,0,0)",
                                        Width = 2.0f
                                    }
                                },
                                // 圆形
                                new ShapeJsonData
                                {
                                    Type = "autoshape",
                                    Name = "Circle",
                                    Box = "12,4,6,6",
                                    SpecialType = "circle",
                                    Fill = new FillJsonData
                                    {
                                        Color = "RGB(100,150,200)",
                                        Opacity = 1.0f
                                    }
                                },
                                // 内容文本框
                                new ShapeJsonData
                                {
                                    Type = "textbox",
                                    Name = "Content",
                                    Box = "2,11,20,8",
                                    HasText = 1,
                                    Text = new List<TextRunJsonData>
                                    {
                                        new TextRunJsonData
                                        {
                                            Content = "这是内容区域。",
                                            Font = "Arial",
                                            FontSize = 14,
                                            FontColor = "RGB(0,0,0)"
                                        },
                                        new TextRunJsonData
                                        {
                                            Content = "可以包含多段文本。",
                                            Font = "Arial",
                                            FontSize = 14,
                                            FontColor = "RGB(64,64,64)"
                                        }
                                    }
                                }
                            }
                        }
                    }
                };

                writer.WriteFromJsonData(jsonData);
                writer.SaveAs(outputPath);

                Console.WriteLine($"成功生成复杂幻灯片: {outputPath}");
            }
        }

        /// <summary>
        /// 示例 4: 从现有 PPTX 读取并修改后保存
        /// </summary>
        public static void Example4_ReadModifyWrite()
        {
            string templatePath = "26xdemo2.pptx";
            string outputPath = "modified_output.pptx";

            // 步骤 1: 读取模板文件为 JSON
            string originalJson = null;
            using (var reader = PowerPointReaderFactory.CreateReader(templatePath, out bool success))
            {
                if (!success)
                {
                    Console.WriteLine("读取模板文件失败");
                    return;
                }

                originalJson = reader.ToJson();
                Console.WriteLine("成功读取模板文件");
            }

            // 步骤 2: 解析 JSON（这里简化处理，实际应该解析并修改）
            var converter = new Core.Converters.JsonToVstoConverter();
            var presentationData = converter.ParseJson(originalJson);

            // 步骤 3: 修改数据（添加新幻灯片）
            if (presentationData.ContentSlides == null)
            {
                presentationData.ContentSlides = new List<SlideJsonData>();
            }

            presentationData.ContentSlides.Add(new SlideJsonData
            {
                PageNumber = presentationData.ContentSlides.Count + 1,
                Title = "新增幻灯片",
                Shapes = new List<ShapeJsonData>
                {
                    new ShapeJsonData
                    {
                        Type = "textbox",
                        Name = "NewSlideTitle",
                        Box = "2,2,20,3",
                        HasText = 1,
                        Text = new List<TextRunJsonData>
                        {
                            new TextRunJsonData
                            {
                                Content = "这是新增的幻灯片",
                                Font = "Arial",
                                FontSize = 20,
                                FontColor = "RGB(0,0,0)"
                            }
                        }
                    }
                }
            });

            // 步骤 4: 使用 VSTO 写入
            using (var writer = PowerPointWriterFactory.CreateWriter())
            {
                writer.OpenFromTemplate(templatePath);
                writer.ClearAllContentSlides();
                writer.WriteFromJsonData(presentationData);
                writer.SaveAs(outputPath);

                Console.WriteLine($"成功生成修改后的文件: {outputPath}");
            }
        }

        /// <summary>
        /// 示例 5: 批量生成多个文件
        /// </summary>
        public static void Example5_BatchGeneration()
        {
            string templatePath = "26xdemo2.pptx";
            string[] outputPaths = { "output1.pptx", "output2.pptx", "output3.pptx" };

            for (int i = 0; i < outputPaths.Length; i++)
            {
                using (var writer = PowerPointWriterFactory.CreateWriter())
                {
                    writer.OpenFromTemplate(templatePath);
                    writer.ClearAllContentSlides();

                    var jsonData = new PresentationJsonData
                    {
                        ContentSlides = new List<SlideJsonData>
                        {
                            new SlideJsonData
                            {
                                PageNumber = 1,
                                Title = $"文件 {i + 1}",
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
                                                Content = $"这是第 {i + 1} 个文件",
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

                    writer.WriteFromJsonData(jsonData);
                    writer.SaveAs(outputPaths[i]);

                    Console.WriteLine($"成功生成: {outputPaths[i]}");
                }
            }
        }

        /// <summary>
        /// 运行所有示例入口方法（不作为程序主入口点）。
        /// 调用示例：在调试时从 Program.Main 中手动调用。
        /// </summary>
        public static void RunExamples(string[] args)
        {
            Console.WriteLine("========== PowerPointWriter 使用示例 ==========\n");

            // 检查模板文件是否存在
            if (!File.Exists("26xdemo2.pptx"))
            {
                Console.WriteLine("错误: 模板文件 26xdemo2.pptx 不存在");
                Console.WriteLine("请确保模板文件在当前目录中");
                return;
            }

            try
            {
                Console.WriteLine("示例 1: 简单使用（需要 content.json 文件）");
                // Example1_SimpleUsage();

                Console.WriteLine("\n示例 2: 创建简单幻灯片");
                Example2_CreateSimpleSlide();

                Console.WriteLine("\n示例 3: 创建复杂幻灯片");
                Example3_CreateComplexSlide();

                Console.WriteLine("\n示例 4: 读取-修改-写入");
                Example4_ReadModifyWrite();

                Console.WriteLine("\n示例 5: 批量生成");
                Example5_BatchGeneration();

                Console.WriteLine("\n========== 所有示例完成 ==========");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"运行示例时出错: {ex.Message}");
                Console.WriteLine($"堆栈跟踪: {ex.StackTrace}");
            }
        }
    }
}

