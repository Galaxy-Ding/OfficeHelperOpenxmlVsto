using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using OfficeHelperOpenXml.Api;
using OfficeHelperOpenXml.Api.Excel;
using OfficeHelperOpenXml.Api.Word;
using OfficeHelperOpenXml.Elements;
using OfficeHelperOpenXml.Components;
using OfficeHelperOpenXml.Models;

namespace OfficeHelperOpenXml.Test
{
    class Program
    {
        // 使用相对路径，优先查找当前项目目录下的测试文件
        // 如果不存在，可以尝试查找原始路径
        static string GetTemplatesPath()
        {
            // 首先尝试当前项目目录
            var currentDir = Directory.GetCurrentDirectory();
            var projectRoot = Path.GetFullPath(Path.Combine(currentDir, "..", ".."));
            var localTemplates = Path.Combine(projectRoot, "templates");
            if (Directory.Exists(localTemplates))
                return localTemplates;
            
            // 尝试原始路径（如果存在）
            var originalPath = "D:/pythonf/office_helper/OfficeHelper/examples/templates";
            if (Directory.Exists(originalPath))
                return originalPath;
            
            // 如果都不存在，返回当前目录
            return currentDir;
        }
        
        static string TemplatesPath => GetTemplatesPath();

        static void Main(string[] args)
        {
            Console.WriteLine("╔══════════════════════════════════════════════════════════════╗");
            Console.WriteLine("║         OfficeHelperOpenXml - Phase 6: Full Testing          ║");
            Console.WriteLine("╚══════════════════════════════════════════════════════════════╝");
            Console.WriteLine($"Version: {OfficeHelperWrapper.GetVersion()}\n");

            if (args.Length > 0)
            {
                switch (args[0])
                {
                    case "--ppt": TestAllPPT(); return;
                    case "--excel": TestAllExcel(); return;
                    case "--word": TestAllWord(); return;
                    case "--write": TestWriteFunctions(); return;
                    case "--perf": TestPerformance(); return;
                    case "--full": RunFullTest(); return;
                    case "--json": TestJsonOutput(); return;
                    case "--analyze": AnalyzeUserFiles(); return;
                    case "--conn": AnalyzeConnection(); return;
                    case "--textbox-props": TestTextboxProperties(); return;
                    default:
                        if (File.Exists(args[0])) { TestSingleFile(args[0]); return; }
                        break;
                }
            }
            RunFullTest();
        }

        static void RunFullTest()
        {
            var sw = Stopwatch.StartNew();
            int passed = 0, failed = 0;

            Console.WriteLine("\n═══════════════════ 1. PPT READ TESTS ═══════════════════\n");
            var pptResults = TestAllPPT();
            passed += pptResults.Item1; failed += pptResults.Item2;

            Console.WriteLine("\n═══════════════════ 2. EXCEL READ TESTS ═══════════════════\n");
            var excelResults = TestAllExcel();
            passed += excelResults.Item1; failed += excelResults.Item2;

            Console.WriteLine("\n═══════════════════ 3. WORD READ TESTS ═══════════════════\n");
            var wordResults = TestAllWord();
            passed += wordResults.Item1; failed += wordResults.Item2;

            Console.WriteLine("\n═══════════════════ 4. WRITE FUNCTION TESTS ═══════════════════\n");
            var writeResults = TestWriteFunctions();
            passed += writeResults.Item1; failed += writeResults.Item2;

            Console.WriteLine("\n═══════════════════ 5. PERFORMANCE TESTS ═══════════════════\n");
            TestPerformance();

            sw.Stop();
            Console.WriteLine("\n╔══════════════════════════════════════════════════════════════╗");
            Console.WriteLine($"║  TOTAL: {passed} PASSED, {failed} FAILED   |   Time: {sw.ElapsedMilliseconds}ms");
            Console.WriteLine("╚══════════════════════════════════════════════════════════════╝");
        }

        static (int, int) TestAllPPT()
        {
            int passed = 0, failed = 0;
            var pptFiles = new[] { "26xdemo1.pptx", "group.pptx", "picture.pptx", "table.pptx" };
            
            foreach (var file in pptFiles)
            {
                var path = Path.Combine(TemplatesPath, file);
                if (!File.Exists(path)) { Console.WriteLine($"  [SKIP] {file} not found"); continue; }
                
                try
                {
                    var sw = Stopwatch.StartNew();
                    using (var reader = new PowerPointReader())
                    {
                        reader.Load(path);
                        var json = reader.ToJson();
                        sw.Stop();
                        var info = reader.PresentationInfo;
                        Console.WriteLine($"  [PASS] {file}");
                        Console.WriteLine($"         Slides: {info.SlideCount}, Elements: {reader.GetAllElements().Count}, JSON: {json.Length} chars, Time: {sw.ElapsedMilliseconds}ms");
                        passed++;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"  [FAIL] {file}: {ex.Message}");
                    failed++;
                }
            }
            return (passed, failed);
        }

        static (int, int) TestAllExcel()
        {
            int passed = 0, failed = 0;
            var excelFiles = new[] { "26xdemo1.xlsx", "first.xlsx", "template_mozumingxi.xlsx" };
            
            foreach (var file in excelFiles)
            {
                var path = Path.Combine(TemplatesPath, file);
                if (!File.Exists(path)) { Console.WriteLine($"  [SKIP] {file} not found"); continue; }
                
                try
                {
                    var sw = Stopwatch.StartNew();
                    using (var reader = new ExcelReader())
                    {
                        reader.Load(path);
                        var sheets = reader.GetSheetNames();
                        var json = reader.ToJson();
                        sw.Stop();
                        Console.WriteLine($"  [PASS] {file}");
                        Console.WriteLine($"         Sheets: {sheets.Count} [{string.Join(", ", sheets.Take(3))}{(sheets.Count > 3 ? "..." : "")}], JSON: {json.Length} chars, Time: {sw.ElapsedMilliseconds}ms");
                        passed++;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"  [FAIL] {file}: {ex.Message}");
                    failed++;
                }
            }
            return (passed, failed);
        }

        static (int, int) TestAllWord()
        {
            int passed = 0, failed = 0;
            var wordFiles = new[] { "26xdemo1.docx" };
            
            foreach (var file in wordFiles)
            {
                var path = Path.Combine(TemplatesPath, file);
                if (!File.Exists(path)) { Console.WriteLine($"  [SKIP] {file} not found"); continue; }
                
                try
                {
                    var sw = Stopwatch.StartNew();
                    using (var reader = new WordReader())
                    {
                        reader.Load(path);
                        var paras = reader.GetParagraphCount();
                        var tables = reader.GetTableCount();
                        var json = reader.ToJson();
                        sw.Stop();
                        Console.WriteLine($"  [PASS] {file}");
                        Console.WriteLine($"         Paragraphs: {paras}, Tables: {tables}, JSON: {json.Length} chars, Time: {sw.ElapsedMilliseconds}ms");
                        passed++;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"  [FAIL] {file}: {ex.Message}");
                    failed++;
                }
            }
            return (passed, failed);
        }

        static (int, int) TestWriteFunctions()
        {
            int passed = 0, failed = 0;
            string tempDir = Path.GetTempPath();

            // Test 1: PPT Write - SKIPPED (PowerPointWriter has been removed)
            Console.WriteLine("  Testing PPT Write...");
            Console.WriteLine("  [SKIP] PPT Write - Feature removed (JSON to PPTX conversion removed)");

            // Test 2: Excel Write
            Console.WriteLine("  Testing Excel Write...");
            try
            {
                string path = Path.Combine(tempDir, "test_write.xlsx");
                if (File.Exists(path)) File.Delete(path);
                using (var writer = new ExcelWriter())
                {
                    writer.OpenOrCreate(path);
                    var data = new List<Dictionary<string, object>>
                    {
                        new Dictionary<string, object> { {"Col1", "A"}, {"Col2", "B"} },
                        new Dictionary<string, object> { {"Col1", "C"}, {"Col2", "D"} }
                    };
                    writer.WriteData("Sheet1", data);
                    writer.Save();
                }
                // Verify
                using (var reader = new ExcelReader())
                {
                    reader.Load(path);
                    var rows = reader.GetSheetData("Sheet1");
                    if (rows.Count == 2) { Console.WriteLine("  [PASS] Excel Write"); passed++; }
                    else { Console.WriteLine($"  [FAIL] Excel Write: expected 2 rows, got {rows.Count}"); failed++; }
                }
            }
            catch (Exception ex) { Console.WriteLine($"  [FAIL] Excel Write: {ex.Message}"); failed++; }

            // Test 3: Word Write
            Console.WriteLine("  Testing Word Write...");
            try
            {
                string path = Path.Combine(tempDir, "test_write.docx");
                if (File.Exists(path)) File.Delete(path);
                using (var writer = new WordWriter())
                {
                    writer.OpenOrCreate(path);
                    writer.AddHeading("Title", 1);
                    writer.AddParagraph("Paragraph 1");
                    writer.AddParagraph("Paragraph 2", true, false, 14);
                    writer.AddTable(new List<List<string>> { new List<string> { "A", "B" }, new List<string> { "1", "2" } });
                    writer.Save();
                }
                // Verify
                using (var reader = new WordReader())
                {
                    reader.Load(path);
                    if (reader.GetParagraphCount() >= 3 && reader.GetTableCount() >= 1) { Console.WriteLine("  [PASS] Word Write"); passed++; }
                    else { Console.WriteLine("  [FAIL] Word Write: verification failed"); failed++; }
                }
            }
            catch (Exception ex) { Console.WriteLine($"  [FAIL] Word Write: {ex.Message}"); failed++; }

            return (passed, failed);
        }

        static void TestPerformance()
        {
            Console.WriteLine("  Performance comparison (5 iterations each):\n");
            var files = new[]
            {
                ("26xdemo1.pptx", "PPT Small"),
                ("group.pptx", "PPT Group"),
                ("table.pptx", "PPT Table"),
            };

            foreach (var (file, desc) in files)
            {
                var path = Path.Combine(TemplatesPath, file);
                if (!File.Exists(path)) continue;

                var times = new List<long>();
                for (int i = 0; i < 5; i++)
                {
                    var sw = Stopwatch.StartNew();
                    using (var reader = new PowerPointReader())
                    {
                        reader.Load(path);
                        reader.ToJson();
                    }
                    sw.Stop();
                    times.Add(sw.ElapsedMilliseconds);
                }
                Console.WriteLine($"  {desc,-15} : First={times[0]}ms, Avg(2-5)={times.Skip(1).Average():F1}ms, Min={times.Min()}ms");
            }

            // Excel performance
            var excelPath = Path.Combine(TemplatesPath, "template_mozumingxi.xlsx");
            if (File.Exists(excelPath))
            {
                var times = new List<long>();
                for (int i = 0; i < 5; i++)
                {
                    var sw = Stopwatch.StartNew();
                    using (var reader = new ExcelReader())
                    {
                        reader.Load(excelPath);
                        reader.ToJson();
                    }
                    sw.Stop();
                    times.Add(sw.ElapsedMilliseconds);
                }
                Console.WriteLine($"  {"Excel Large",-15} : First={times[0]}ms, Avg(2-5)={times.Skip(1).Average():F1}ms, Min={times.Min()}ms");
            }
        }

        static void TestSingleFile(string path)
        {
            Console.WriteLine($"Testing: {path}\n");
            var ext = Path.GetExtension(path).ToLower();
            var sw = Stopwatch.StartNew();

            if (ext == ".pptx")
            {
                using (var reader = new PowerPointReader())
                {
                    reader.Load(path);
                    var json = reader.ToJson();
                    sw.Stop();
                    Console.WriteLine($"Slides: {reader.PresentationInfo.SlideCount}");
                    Console.WriteLine($"Elements: {reader.GetAllElements().Count}");
                    Console.WriteLine($"JSON length: {json.Length}");
                    Console.WriteLine($"Time: {sw.ElapsedMilliseconds}ms");

                    // Save full JSON to file
                    var outputPath = Path.ChangeExtension(path, ".json");
                    File.WriteAllText(outputPath, json);
                    Console.WriteLine($"Full JSON saved to: {outputPath}");

                    Console.WriteLine($"\nJSON preview:\n{(json.Length > 2000 ? json.Substring(0, 2000) + "..." : json)}");
                }
            }
            else if (ext == ".xlsx")
            {
                using (var reader = new ExcelReader())
                {
                    reader.Load(path);
                    var json = reader.ToJson();
                    sw.Stop();
                    Console.WriteLine($"Sheets: {string.Join(", ", reader.GetSheetNames())}");
                    Console.WriteLine($"JSON length: {json.Length}");
                    Console.WriteLine($"Time: {sw.ElapsedMilliseconds}ms");
                    Console.WriteLine($"\nJSON preview:\n{(json.Length > 2000 ? json.Substring(0, 2000) + "..." : json)}");
                }
            }
            else if (ext == ".docx")
            {
                using (var reader = new WordReader())
                {
                    reader.Load(path);
                    var json = reader.ToJson();
                    sw.Stop();
                    Console.WriteLine($"Paragraphs: {reader.GetParagraphCount()}");
                    Console.WriteLine($"Tables: {reader.GetTableCount()}");
                    Console.WriteLine($"JSON length: {json.Length}");
                    Console.WriteLine($"Time: {sw.ElapsedMilliseconds}ms");
                    Console.WriteLine($"\nJSON preview:\n{(json.Length > 2000 ? json.Substring(0, 2000) + "..." : json)}");
                }
            }
        }

        static void TestJsonOutput()
        {
            Console.WriteLine("\n═══════════════════ JSON FORMAT TEST ═══════════════════\n");
            
            var testFiles = new[] { "picture.pptx", "26xdemo1.pptx", "group.pptx" };
            
            foreach (var file in testFiles)
            {
                string pptPath = Path.Combine(TemplatesPath, file);
                // 输出到当前项目目录
                var outputDir = Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory(), "..", ".."));
                string outputPath = Path.Combine(outputDir, $"test_output_{Path.GetFileNameWithoutExtension(file)}.json");
                
                if (!File.Exists(pptPath))
                {
                    Console.WriteLine($"File not found: {pptPath}");
                    continue;
                }
                
                Console.WriteLine($"\n--- Testing {file} ---");
                
                using (var reader = new PowerPointReader())
                {
                    reader.Load(pptPath);
                    var json = reader.ToJson();
                    
                    File.WriteAllText(outputPath, json);
                    
                    Console.WriteLine($"✓ JSON saved to: {outputPath}");
                    Console.WriteLine($"✓ JSON length: {json.Length} chars");
                    Console.WriteLine($"✓ Slides: {reader.PresentationInfo.SlideCount}");
                    Console.WriteLine($"✓ Elements: {reader.GetAllElements().Count}");
                    
                    Console.WriteLine($"\n--- JSON Preview (first 1500 chars) ---");
                    Console.WriteLine(json.Substring(0, Math.Min(1500, json.Length)));
                    if (json.Length > 1500) Console.WriteLine("...");
                }
            }
        }
        
        static void AnalyzeUserFiles()
        {
            var files = new[]
            {
                @"D:\pythonf\office_helper\OfficeHelper\examples\templates\方案報告(改造-單機-連線)-版本20251106更新.pptx",
                @"D:\pythonf\office_helper\OfficeHelper\examples\templates\方案報告(新增&再製-單機-連線)-版本20251106更新.pptx",
                @"D:\download\263 CNC7.2螺紋孔小徑檢測機改造253PL-DFM-V3.1 LYG--0406.pptx",
                @"D:\download\263Forging尺寸檢測機再制再制-DFM-V3-20250711.pptx"
            };
            
            AnalyzePPT.AnalyzeFiles(files);
        }
        
        static void AnalyzeConnection()
        {
            var file = @"D:\pythonf\office_helper\OfficeHelper\examples\templates\方案報告(改造-單機-連線)-版本20251106更新.pptx";
            AnalyzeConnectionShape.Analyze(file);
        }

        static void TestTextboxProperties()
        {
            Console.WriteLine("\n═══════════════════ TEXTBOX PROPERTIES TEST ═══════════════════\n");
            
            var testDir = AppDomain.CurrentDomain.BaseDirectory;
            var solutionRoot = Path.GetFullPath(Path.Combine(testDir, "..", "..", "..", ".."));
            var templatePath = Path.Combine(solutionRoot, "OfficeHelperOpenxmVsto.Test", "TestTemplates", "textbox_properties_test_template.pptx");
            
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"测试模板不存在: {templatePath}");
                Console.WriteLine("\n请按照以下步骤创建测试模板:");
                Console.WriteLine("1. 创建包含9个文本框的PPTX文件");
                Console.WriteLine("2. 每个文本框设置不同的属性组合");
                Console.WriteLine("3. 文本框内容格式: {字体名}-({R},{G},{B})-{字号}pt");
                Console.WriteLine("\n详细说明请参考: test_implement.md");
                return;
            }

            Console.WriteLine($"加载测试模板: {templatePath}");
            
            try
            {
                using (var reader = new PowerPointReader())
                {
                    reader.Load(templatePath);
                    var json = reader.ToJson();
                    
                    var outputDir = Path.Combine(solutionRoot, "OfficeHelperOpenxmVsto.Test", "test_output");
                    Directory.CreateDirectory(outputDir);
                    
                    var jsonPath = Path.Combine(outputDir, "textbox_properties_test.json");
                    File.WriteAllText(jsonPath, json);
                    
                    Console.WriteLine($"✓ JSON已保存到: {jsonPath}");
                    Console.WriteLine($"✓ JSON长度: {json.Length} 字符");
                    Console.WriteLine($"✓ 幻灯片数: {reader.PresentationInfo.SlideCount}");
                    Console.WriteLine($"✓ 元素总数: {reader.GetAllElements().Count}");
                    
                    Console.WriteLine("\n提示: 使用 xUnit 测试运行器运行完整测试:");
                    Console.WriteLine("  dotnet test --filter TextboxPropertyCompletenessTest");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"错误: {ex.Message}");
                Console.WriteLine($"堆栈跟踪: {ex.StackTrace}");
            }
        }
    }
}
