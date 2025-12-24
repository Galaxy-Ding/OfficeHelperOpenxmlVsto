using System;
using System.IO;
using System.Linq;
using Xunit;
using OfficeHelperOpenXml.Api;
using OfficeHelperOpenXml.Api.PowerPoint;
using OfficeHelperOpenXml.Models.Json;
using OfficeHelperOpenXml.Utils;

namespace OfficeHelperOpenXml.Test
{
    /// <summary>
    /// PowerPointWriter 端到端集成测试
    /// 阶段三：完整流程测试
    /// </summary>
    public class PowerPointWriterIntegrationTests : IDisposable
    {
        private static string TemplatePath => TestPaths.Template26xdemo2Path;
        private const string TestOutputPath = "test_integration_output.pptx";
        private const string TestJsonPath = "test_integration_data.json";

        public void Dispose()
        {
            // 清理测试输出文件
            try
            {
                if (File.Exists(TestOutputPath))
                    File.Delete(TestOutputPath);
                if (File.Exists(TestJsonPath))
                    File.Delete(TestJsonPath);
            }
            catch { }
        }

        /// <summary>
        /// 端到端测试：读取 → 修改 → 写入 → 验证
        /// </summary>
        [Fact]
        public void TestEndToEnd_ReadModifyWriteVerify()
        {
            if (!File.Exists(TemplatePath))
            {
                Assert.True(false, $"模板文件不存在: {TemplatePath}");
                return;
            }

            if (!VstoHelper.IsPowerPointAvailable())
            {
                Assert.True(false, "PowerPoint 不可用");
                return;
            }

            // 步骤 1: 读取模板文件 → JSON（使用 OpenXML）
            string originalJson = null;
            using (var reader = PowerPointReaderFactory.CreateReader(TemplatePath, out bool success))
            {
                Assert.True(success, "应该能够读取模板文件");
                originalJson = reader.ToJson();
                Assert.NotNull(originalJson);
                Assert.True(originalJson.Length > 0, "JSON 数据不应该为空");
            }

            // 步骤 2: 解析 JSON 数据
            var converter = new Core.Converters.JsonToVstoConverter();
            var presentationData = converter.ParseJson(originalJson);
            Assert.NotNull(presentationData);

            // 步骤 3: 修改 JSON 数据（添加一个简单的文本框）
            if (presentationData.ContentSlides == null)
            {
                presentationData.ContentSlides = new System.Collections.Generic.List<SlideJsonData>();
            }

            if (presentationData.ContentSlides.Count == 0)
            {
                presentationData.ContentSlides.Add(new SlideJsonData
                {
                    PageNumber = 1,
                    Title = "Test Slide"
                });
            }

            var firstSlide = presentationData.ContentSlides[0];
            if (firstSlide.Shapes == null)
            {
                firstSlide.Shapes = new System.Collections.Generic.List<ShapeJsonData>();
            }

            // 添加一个测试文本框
            firstSlide.Shapes.Add(new ShapeJsonData
            {
                Type = "textbox",
                Name = "IntegrationTestTextBox",
                Box = "5,5,10,3",
                HasText = 1,
                Text = new System.Collections.Generic.List<TextRunJsonData>
                {
                    new TextRunJsonData
                    {
                        Content = "Integration Test",
                        Font = "Arial",
                        FontSize = 16,
                        FontColor = "RGB(0,0,0)"
                    }
                }
            });

            // 步骤 4: 使用 VSTO 写入 → 新 PPTX
            using (var writer = PowerPointWriterFactory.CreateWriter())
            {
                Assert.True(writer.OpenFromTemplate(TemplatePath), "应该能够打开模板");
                Assert.True(writer.ClearAllContentSlides(), "应该能够清除内容幻灯片");
                Assert.True(writer.WriteFromJsonData(presentationData), "应该能够写入 JSON 数据");

                if (File.Exists(TestOutputPath))
                {
                    File.Delete(TestOutputPath);
                }
                Assert.True(writer.SaveAs(TestOutputPath), "应该能够保存文件");
                Assert.True(File.Exists(TestOutputPath), "输出文件应该存在");
            }

            // 步骤 5: 读取新 PPTX → JSON（使用 OpenXML）
            using (var reader = PowerPointReaderFactory.CreateReader(TestOutputPath, out bool success))
            {
                Assert.True(success, "应该能够读取生成的 PPTX 文件");
                string newJson = reader.ToJson();
                Assert.NotNull(newJson);
                Assert.True(newJson.Length > 0, "生成的 JSON 数据不应该为空");

                var newPresentationData = converter.ParseJson(newJson);
                Assert.NotNull(newPresentationData);

                // 步骤 6: 对比关键字段
                if (newPresentationData.ContentSlides != null && newPresentationData.ContentSlides.Count > 0)
                {
                    var newFirstSlide = newPresentationData.ContentSlides[0];
                    Assert.NotNull(newFirstSlide.Shapes);
                    Assert.True(newFirstSlide.Shapes != null, "新幻灯片应该有形状列表");

                    // 验证形状数量（至少应该有我们添加的文本框）
                    if (newFirstSlide.Shapes != null)
                    {
                        Assert.True(newFirstSlide.Shapes.Count > 0, "新幻灯片应该至少有一个形状");
                        
                        // 查找我们添加的文本框
                        var testTextBox = newFirstSlide.Shapes.FirstOrDefault(s => 
                            s.Name == "IntegrationTestTextBox" || 
                            (s.Type == "textbox" && s.Text != null && s.Text.Count > 0));
                        
                        Assert.NotNull(testTextBox);
                        Assert.True(testTextBox != null, "应该能找到测试文本框");
                    }
                }
            }
        }

        /// <summary>
        /// 测试模板的母版样式是否正确保留
        /// </summary>
        [Fact]
        public void TestTemplateMasterStylesPreserved()
        {
            if (!File.Exists(TemplatePath))
            {
                Assert.True(false, $"模板文件不存在: {TemplatePath}");
                return;
            }

            if (!VstoHelper.IsPowerPointAvailable())
            {
                Assert.True(false, "PowerPoint 不可用");
                return;
            }

            // 读取模板的母版样式
            using (var reader = PowerPointReaderFactory.CreateReader(TemplatePath, out bool success))
            {
                Assert.True(success);
                var info = reader.PresentationInfo;
                Assert.NotNull(info);
                Assert.NotNull(info.SlideMasterStyles);
                Assert.True(info.SlideMasterStyles != null, "模板应该有母版样式");
            }

            // 写入并保存
            var jsonData = new PresentationJsonData
            {
                ContentSlides = new System.Collections.Generic.List<SlideJsonData>
                {
                    new SlideJsonData
                    {
                        PageNumber = 1,
                        Title = "Style Test"
                    }
                }
            };

            using (var writer = PowerPointWriterFactory.CreateWriter())
            {
                Assert.True(writer.OpenFromTemplate(TemplatePath));
                Assert.True(writer.ClearAllContentSlides());
                Assert.True(writer.WriteFromJsonData(jsonData));

                if (File.Exists(TestOutputPath))
                {
                    File.Delete(TestOutputPath);
                }
                Assert.True(writer.SaveAs(TestOutputPath));
            }

            // 验证生成的文件的母版样式
            using (var reader = PowerPointReaderFactory.CreateReader(TestOutputPath, out bool success))
            {
                Assert.True(success);
                var info = reader.PresentationInfo;
                Assert.NotNull(info);
                // 验证母版样式仍然存在
                Assert.NotNull(info.SlideMasterStyles);
                Assert.True(info.SlideMasterStyles != null, "生成的文件应该有母版样式");
            }
        }

        /// <summary>
        /// 测试空 JSON 数据
        /// </summary>
        [Fact]
        public void TestEmptyJsonData()
        {
            if (!File.Exists(TemplatePath))
            {
                Assert.True(false, $"模板文件不存在: {TemplatePath}");
                return;
            }

            if (!VstoHelper.IsPowerPointAvailable())
            {
                Assert.True(false, "PowerPoint 不可用");
                return;
            }

            var emptyJsonData = new PresentationJsonData
            {
                ContentSlides = new System.Collections.Generic.List<SlideJsonData>()
            };

            using (var writer = PowerPointWriterFactory.CreateWriter())
            {
                Assert.True(writer.OpenFromTemplate(TemplatePath));
                Assert.True(writer.ClearAllContentSlides());
                
                // 空数据应该也能处理（不报错）
                bool success = writer.WriteFromJsonData(emptyJsonData);
                Assert.True(success, "空 JSON 数据应该能够处理");

                if (File.Exists(TestOutputPath))
                {
                    File.Delete(TestOutputPath);
                }
                Assert.True(writer.SaveAs(TestOutputPath));
                Assert.True(File.Exists(TestOutputPath));
            }
        }

        /// <summary>
        /// 测试大量形状（性能测试）
        /// </summary>
        [Fact]
        public void TestLargeNumberOfShapes()
        {
            if (!File.Exists(TemplatePath))
            {
                Assert.True(false, $"模板文件不存在: {TemplatePath}");
                return;
            }

            if (!VstoHelper.IsPowerPointAvailable())
            {
                Assert.True(false, "PowerPoint 不可用");
                return;
            }

            // 创建包含 50 个形状的幻灯片
            var shapes = new System.Collections.Generic.List<ShapeJsonData>();
            for (int i = 0; i < 50; i++)
            {
                shapes.Add(new ShapeJsonData
                {
                    Type = "textbox",
                    Name = $"Shape{i}",
                    Box = $"{i % 5 * 4},{i / 5 * 2},3,1.5",
                    HasText = 1,
                    Text = new System.Collections.Generic.List<TextRunJsonData>
                    {
                        new TextRunJsonData
                        {
                            Content = $"Shape {i}",
                            Font = "Arial",
                            FontSize = 10,
                            FontColor = "RGB(0,0,0)"
                        }
                    }
                });
            }

            var jsonData = new PresentationJsonData
            {
                ContentSlides = new System.Collections.Generic.List<SlideJsonData>
                {
                    new SlideJsonData
                    {
                        PageNumber = 1,
                        Title = "Large Shape Test",
                        Shapes = shapes
                    }
                }
            };

            var startTime = DateTime.Now;

            using (var writer = PowerPointWriterFactory.CreateWriter())
            {
                Assert.True(writer.OpenFromTemplate(TemplatePath));
                Assert.True(writer.ClearAllContentSlides());
                Assert.True(writer.WriteFromJsonData(jsonData));

                if (File.Exists(TestOutputPath))
                {
                    File.Delete(TestOutputPath);
                }
                Assert.True(writer.SaveAs(TestOutputPath));
            }

            var elapsed = DateTime.Now - startTime;
            Console.WriteLine($"写入 50 个形状耗时: {elapsed.TotalSeconds:F2} 秒");

            // 验证文件已创建
            Assert.True(File.Exists(TestOutputPath));
            
            // 验证性能（应该小于 20 秒）
            Assert.True(elapsed.TotalSeconds < 20, $"写入 50 个形状应该在 20 秒内完成，实际耗时: {elapsed.TotalSeconds:F2} 秒");
        }

        /// <summary>
        /// 测试复杂样式（渐变填充、图片填充等）
        /// </summary>
        [Fact]
        public void TestComplexStyles()
        {
            if (!File.Exists(TemplatePath))
            {
                Assert.True(false, $"模板文件不存在: {TemplatePath}");
                return;
            }

            if (!VstoHelper.IsPowerPointAvailable())
            {
                Assert.True(false, "PowerPoint 不可用");
                return;
            }

            var jsonData = new PresentationJsonData
            {
                ContentSlides = new System.Collections.Generic.List<SlideJsonData>
                {
                    new SlideJsonData
                    {
                        PageNumber = 1,
                        Title = "Complex Style Test",
                        Shapes = new System.Collections.Generic.List<ShapeJsonData>
                        {
                            // 带填充和线条的形状
                            new ShapeJsonData
                            {
                                Type = "autoshape",
                                Name = "StyledShape",
                                Box = "2,2,8,6",
                                SpecialType = "rectangle",
                                Fill = new FillJsonData
                                {
                                    Color = "RGB(100,150,200)",
                                    Opacity = 0.8f
                                },
                                Line = new LineJsonData
                                {
                                    HasOutline = 1,
                                    Color = "RGB(0,0,0)",
                                    Width = 2.0f
                                },
                                Shadow = new ShadowJsonData
                                {
                                    HasShadow = 1,
                                    Color = "RGB(128,128,128)",
                                    Blur = 5.0f,
                                    OffsetX = 2.0f,
                                    OffsetY = 2.0f
                                }
                            }
                        }
                    }
                }
            };

            using (var writer = PowerPointWriterFactory.CreateWriter())
            {
                Assert.True(writer.OpenFromTemplate(TemplatePath));
                Assert.True(writer.ClearAllContentSlides());
                Assert.True(writer.WriteFromJsonData(jsonData));

                if (File.Exists(TestOutputPath))
                {
                    File.Delete(TestOutputPath);
                }
                Assert.True(writer.SaveAs(TestOutputPath));
                Assert.True(File.Exists(TestOutputPath));
            }
        }
    }
}

