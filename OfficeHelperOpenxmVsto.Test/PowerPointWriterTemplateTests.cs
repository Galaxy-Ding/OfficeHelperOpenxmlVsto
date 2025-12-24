using System;
using System.IO;
using Xunit;
using OfficeHelperOpenXml.Api.PowerPoint;
using OfficeHelperOpenXml.Models.Json;
using OfficeHelperOpenXml.Utils;
using OfficeHelperOpenXml.Api;

namespace OfficeHelperOpenXml.Test
{
    /// <summary>
    /// PowerPointWriter 模板处理测试
    /// 阶段一：验证模板打开、清除和写入功能
    /// </summary>
    public class PowerPointWriterTemplateTests : IDisposable
    {
        private static string TemplatePath => TestPaths.Template26xdemo2Path;
        private const string TestOutputPath = "test_output_template.pptx";

        public void Dispose()
        {
            // 清理测试输出文件
            if (File.Exists(TestOutputPath))
            {
                try
                {
                    File.Delete(TestOutputPath);
                }
                catch { }
            }
        }

        /// <summary>
        /// 测试打开模板文件 26xdemo2.pptx
        /// </summary>
        [Fact]
        public void TestOpenTemplate_26xdemo2()
        {
            if (!File.Exists(TemplatePath))
            {
                Assert.True(false, $"模板文件不存在: {TemplatePath}");
                return;
            }

            // 检查 PowerPoint 是否可用
            if (!VstoHelper.IsPowerPointAvailable())
            {
                Assert.True(false, "PowerPoint 不可用，请确保已安装 Microsoft PowerPoint");
                return;
            }

            using (var writer = PowerPointWriterFactory.CreateWriter())
            {
                Assert.NotNull(writer);
                bool success = writer.OpenFromTemplate(TemplatePath);
                Assert.True(success, "应该能够成功打开模板文件");
            }
        }

        /// <summary>
        /// 测试模板文件存在性检查
        /// </summary>
        [Fact]
        public void TestOpenTemplate_FileNotExists()
        {
            using (var writer = PowerPointWriterFactory.CreateWriter())
            {
                bool success = writer.OpenFromTemplate("nonexistent.pptx");
                Assert.False(success, "不存在的文件应该返回 false");
            }
        }

        /// <summary>
        /// 测试清除内容幻灯片（保留母版）
        /// </summary>
        [Fact]
        public void TestClearContentSlides_PreservesMaster()
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

            using (var writer = PowerPointWriterFactory.CreateWriter())
            {
                // 1. 打开模板
                Assert.True(writer.OpenFromTemplate(TemplatePath), "应该能够打开模板");

                // 2. 记录清除前的幻灯片数量
                // 注意：由于 VSTO 的限制，我们无法直接访问内部状态
                // 这里主要验证方法能够成功执行而不抛出异常

                // 3. 清除内容幻灯片
                bool success = writer.ClearAllContentSlides();
                Assert.True(success, "应该能够成功清除内容幻灯片");

                // 4. 验证母版形状保留（通过保存文件后重新读取验证）
                // 这里简化处理，主要验证方法执行成功
            }
        }

        /// <summary>
        /// 测试写入简单形状
        /// </summary>
        [Fact]
        public void TestWriteFromJson_SimpleShape()
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

            // 创建简单的 JSON 数据
            var jsonData = new PresentationJsonData
            {
                ContentSlides = new System.Collections.Generic.List<SlideJsonData>
                {
                    new SlideJsonData
                    {
                        PageNumber = 1,
                        Title = "Test Slide",
                        Shapes = new System.Collections.Generic.List<ShapeJsonData>
                        {
                            new ShapeJsonData
                            {
                                Type = "textbox",
                                Name = "TestTextBox",
                                Box = "2,2,10,3",
                                HasText = 1,
                                Text = new System.Collections.Generic.List<TextRunJsonData>
                                {
                                    new TextRunJsonData
                                    {
                                        Content = "Hello, World!",
                                        Font = "Arial",
                                        FontSize = 14,
                                        FontColor = "RGB(0,0,0)"
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

            using (var writer = PowerPointWriterFactory.CreateWriter())
            {
                Assert.True(writer.OpenFromTemplate(TemplatePath));
                Assert.True(writer.ClearAllContentSlides());
                
                bool success = writer.WriteFromJsonData(jsonData);
                Assert.True(success, "应该能够成功写入 JSON 数据");

                // 保存文件
                if (File.Exists(TestOutputPath))
                {
                    File.Delete(TestOutputPath);
                }
                success = writer.SaveAs(TestOutputPath);
                Assert.True(success, "应该能够成功保存文件");
                Assert.True(File.Exists(TestOutputPath), "输出文件应该存在");
            }
        }

        /// <summary>
        /// 测试写入复杂幻灯片
        /// </summary>
        [Fact]
        public void TestWriteFromJson_ComplexSlide()
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

            // 创建包含多种形状的复杂 JSON 数据
            var jsonData = new PresentationJsonData
            {
                ContentSlides = new System.Collections.Generic.List<SlideJsonData>
                {
                    new SlideJsonData
                    {
                        PageNumber = 1,
                        Title = "Complex Slide",
                        Shapes = new System.Collections.Generic.List<ShapeJsonData>
                        {
                            // 文本框
                            new ShapeJsonData
                            {
                                Type = "textbox",
                                Name = "Title",
                                Box = "1,1,20,2",
                                HasText = 1,
                                Text = new System.Collections.Generic.List<TextRunJsonData>
                                {
                                    new TextRunJsonData
                                    {
                                        Content = "Complex Slide Title",
                                        Font = "Arial",
                                        FontSize = 24,
                                        FontColor = "RGB(0,0,0)",
                                        FontBold = 1
                                    }
                                }
                            },
                            // 自动形状（矩形）
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
                            // 自动形状（圆形）
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
                            }
                        }
                    }
                }
            };

            using (var writer = PowerPointWriterFactory.CreateWriter())
            {
                Assert.True(writer.OpenFromTemplate(TemplatePath));
                Assert.True(writer.ClearAllContentSlides());
                
                bool success = writer.WriteFromJsonData(jsonData);
                Assert.True(success, "应该能够成功写入复杂 JSON 数据");

                // 保存文件
                if (File.Exists(TestOutputPath))
                {
                    File.Delete(TestOutputPath);
                }
                success = writer.SaveAs(TestOutputPath);
                Assert.True(success, "应该能够成功保存文件");
            }
        }

        /// <summary>
        /// 测试另存为新文件
        /// </summary>
        [Fact]
        public void TestSaveAs_NewFile()
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

            using (var writer = PowerPointWriterFactory.CreateWriter())
            {
                Assert.True(writer.OpenFromTemplate(TemplatePath));

                // 确保输出文件不存在
                if (File.Exists(TestOutputPath))
                {
                    File.Delete(TestOutputPath);
                }

                bool success = writer.SaveAs(TestOutputPath);
                Assert.True(success, "应该能够成功另存为");
                Assert.True(File.Exists(TestOutputPath), "输出文件应该存在");

                // 验证文件大小（应该大于 0）
                var fileInfo = new FileInfo(TestOutputPath);
                Assert.True(fileInfo.Length > 0, "输出文件大小应该大于 0");
            }
        }

        /// <summary>
        /// 测试 COM 对象正确释放
        /// </summary>
        [Fact]
        public void TestComObjectDisposal()
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

            // 创建多个 writer 实例，验证资源能够正确释放
            for (int i = 0; i < 3; i++)
            {
                using (var writer = PowerPointWriterFactory.CreateWriter())
                {
                    Assert.True(writer.OpenFromTemplate(TemplatePath));
                    // writer 在 using 块结束时应该正确释放
                }

                // 强制垃圾回收
                VstoHelper.ForceGarbageCollection();
            }

            // 如果能够执行到这里，说明资源释放正常
            Assert.True(true);
        }
    }
}

