using System;
using System.IO;
using OfficeHelperOpenXml.Api;
using OfficeHelperOpenXml.Elements;
using Xunit;
using Xunit.Abstractions;

namespace OfficeHelperOpenXml.Test
{
    /// <summary>
    /// 使用 OfficeHelperOpenxmVsto 将 textbox.pptx 转换为 JSON，并输出基础分析结果。
    /// 对应用户计划中的步骤：
    /// 1) 使用指定的 PPTX 文件；
    /// 2) 生成 JSON；
    /// 3) 利用测试项目进行验证和分析输出。
    /// </summary>
    public class TextboxPptxJsonConversionTest
    {
        private readonly ITestOutputHelper _output;

        public TextboxPptxJsonConversionTest(ITestOutputHelper output)
        {
            _output = output;

            // 确保输出目录存在
            Directory.CreateDirectory(TestPaths.TestOutputDir);
            Directory.CreateDirectory(TestPaths.TestReportsDir);
        }

        /// <summary>
        /// 步骤 1 & 3: 利用 OfficeHelperOpenxmVsto (PowerPointReader) 将 textbox.pptx 转换为 JSON
        /// </summary>
        [Fact]
        public void Convert_TextboxPptx_To_Json_Success()
        {
            var pptPath = TestPaths.TextboxPptxPath;
            var jsonPath = TestPaths.TextboxJsonOutputPath;

            Assert.True(File.Exists(pptPath), $"测试 PPTX 不存在: {pptPath}");

            using (var reader = new PowerPointReader())
            {
                reader.Load(pptPath);
                var json = reader.ToJson();

                File.WriteAllText(jsonPath, json);

                _output.WriteLine($"✅ 已从 textbox.pptx 生成 JSON: {jsonPath}");
                _output.WriteLine($"   JSON 长度: {json.Length} 字符");
                _output.WriteLine($"   幻灯片数量: {reader.PresentationInfo?.SlideCount}");
                _output.WriteLine($"   元素总数: {reader.GetAllElements().Count}");

                Assert.True(json.Length > 0, "生成的 JSON 不应为空");
                Assert.True(File.Exists(jsonPath), "JSON 输出文件应该存在");
            }
        }

        /// <summary>
        /// 步骤 4: 对 JSON / PPTX 做一个简单的属性分析并输出 Markdown 报告
        /// （详细的属性逐项验证由 TextboxPropertyCompletenessTest 负责）
        /// </summary>
        [Fact]
        public void Generate_TextboxPptx_Basic_Analysis_Report()
        {
            var pptPath = TestPaths.TextboxPptxPath;
            var reportPath = Path.Combine(TestPaths.TestReportsDir, "textbox_basic_analysis.md");

            Assert.True(File.Exists(pptPath), $"测试 PPTX 不存在: {pptPath}");

            using (var reader = new PowerPointReader())
            {
                reader.Load(pptPath);
                var info = reader.PresentationInfo;
                var elements = reader.GetAllElements();

                int textboxCount = 0;
                foreach (var el in elements)
                {
                    if (el is TextBoxElement)
                    {
                        textboxCount++;
                    }
                }

                using (var sw = new StreamWriter(reportPath, false))
                {
                    sw.WriteLine("# textbox.pptx 基础分析报告");
                    sw.WriteLine();
                    sw.WriteLine($"- **文件路径**: `{pptPath}`");
                    sw.WriteLine($"- **幻灯片数量**: {info?.SlideCount}");
                    sw.WriteLine($"- **页面宽度**: {info?.SlideWidth}");
                    sw.WriteLine($"- **页面高度**: {info?.SlideHeight}");
                    sw.WriteLine($"- **元素总数**: {elements.Count}");
                    sw.WriteLine($"- **文本框数量**: {textboxCount}");
                    sw.WriteLine();
                    sw.WriteLine("> 详细的文本框属性覆盖率和字段级验证请参见 `TextboxPropertyCompletenessTest` 生成的报告。");
                }

                _output.WriteLine($"✅ 已生成基础分析报告: {reportPath}");
            }
        }
    }
}


