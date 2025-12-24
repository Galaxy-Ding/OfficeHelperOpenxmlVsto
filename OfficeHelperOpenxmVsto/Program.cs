using System;
using System.IO;
using System.Collections.Generic;
using OfficeHelperOpenXml.Api;
using OfficeHelperOpenXml.Api.Excel;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace OfficeHelperOpenXml
{
    /// <summary>
    /// OfficeHelperOpenXml 主程序入口
    /// 用于直接调试 PowerPoint/Excel 分析功能
    /// 参考 D:\pythonf\office_helper\OfficeHelper\Program.cs 的调用方式
    /// </summary>
    class Program
    {
        /// <summary>
        /// 主入口点
        /// 必须标记为 [STAThread] 以支持 PowerPoint COM 操作
        /// PowerPoint COM 对象需要在单线程单元 (STA) 中运行
        /// </summary>
        /// <param name="args">命令行参数</param>
        /// <returns>退出代码</returns>
        [STAThread]
        static int Main(string[] args)
        {
            try
            {
                // Check for command-line arguments
                if (args.Length > 0)
                {
                    return ParseCommandLineArguments(args);
                }

                // 显示欢迎信息
                Console.WriteLine("========================================");
                Console.WriteLine("  OfficeHelperOpenXml - 调试工具");
                Console.WriteLine("  基于 OpenXML SDK 的 Office 文件分析");
                Console.WriteLine("========================================");
                Console.WriteLine();

                // ============================================
                // 测试区域 - 根据需要取消注释相应的测试代码
                // ============================================

                // 测试 1: PowerPoint 文件分析
                //string pptPath = @"D:\pythonf\c_sharp_project\OfficeHelperOpenxmVsto\test_ppt\textboxFontMulti.pptx";
                //string outputJsonPath = @"D:\pythonf\c_sharp_project\OfficeHelperOpenxmVsto\test_ppt\textboxFontMulti.json";
                //return ProcessPowerPoint(pptPath, outputJsonPath);

                // 测试 2: Excel 文件分析
                //string excelPath = @"D:\test\sample.xlsx";
                //string excelOutputPath = @"D:\test\output_excel.json";
                //return ProcessExcel(excelPath, excelOutputPath);

                // 测试 3: 从 JSON 恢复 PowerPoint (使用新的转换器)
                //string jsonPath = @"D:\pythonf\office_helper\OfficeHelper\examples\templates\textbox.json";
                //string outputPptPath = @"D:\pythonf\office_helper\OfficeHelper\examples\templates\textbox_json.pptx";
                //return CreatePPTFromJson(jsonPath, outputPptPath);

                // ============================================
                // 测试 1: PowerPoint 文件分析
                // ============================================
                //string pptPath = @"D:\pythonf\c_sharp_project\OfficeHelperOpenxmVsto\textbox.pptx";
                //string outputJsonPath = @"D:\pythonf\c_sharp_project\OfficeHelperOpenxmVsto\outputTextBox.json";
                //return ProcessPowerPoint(pptPath, outputJsonPath);
                // ============================================

                // ============================================
                // 测试: 从 JSON 创建 PowerPoint 文件
                // ============================================
                //获取工作区根目录（向上两级从 bin/ Debug 或 bin/ Release 到项目根目录，再向上到解决方案根目录）
                //string workspaceRoot = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", ".."));
                //string jsonPath = Path.Combine(workspaceRoot, "outputTextBox.json");
                //string templatePath = Path.Combine(workspaceRoot, "26xdemo2.pptx");
                //string outputPptPath = Path.Combine(workspaceRoot, "json_textbox.pptx");
                string jsonPath = @"D:\pythonf\c_sharp_project\OfficeHelperOpenxmVsto\test_ppt\textboxFontMulti.json";
                string templatePath = @"D:\pythonf\c_sharp_project\OfficeHelperOpenxmVsto\26xdemo2.pptx";
                string outputPptPath = @"D:\pythonf\c_sharp_project\OfficeHelperOpenxmVsto\test_ppt\textboxFontMulti_json.pptx";

                // 如果找不到，尝试使用绝对路径
                if (!File.Exists(jsonPath))
                {
                    jsonPath = @"D:\pythonf\office_helper\OfficeHelper\examples\templates\textbox.json";
                }
                if (!File.Exists(templatePath))
                {
                    templatePath = @"D:\pythonf\c_sharp_project\OfficeHelperOpenxmVsto\26xdemo2.pptx";
                }

                return CreatePPTFromJson(jsonPath, templatePath, outputPptPath);
                 //============================================

                 //If no test is enabled, show usage instructions
                ShowUsage();
                return 0;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ 程序执行出错: {ex.Message}");
                Console.WriteLine($"错误详情: {ex.StackTrace}");
                return 1;
            }
        }

        /// <summary>
        /// Parses command-line arguments and executes the appropriate action
        /// </summary>
        /// <param name="args">Command-line arguments</param>
        /// <returns>Exit code</returns>
        private static int ParseCommandLineArguments(string[] args)
        {
            // Check for help flag
            if (args.Length == 1 && (args[0] == "--help" || args[0] == "-h" || args[0] == "/?" || args[0] == "help"))
            {
                ShowCommandLineHelp();
                return 0;
            }

            // Parse mode, input, output, and template arguments
            string mode = null;
            string inputPath = null;
            string outputPath = null;
            string templatePath = null;

            for (int i = 0; i < args.Length; i++)
            {
                if ((args[i] == "--mode" || args[i] == "-m") && i + 1 < args.Length)
                {
                    mode = args[i + 1];
                    i++; // Skip next argument
                }
                else if ((args[i] == "--input" || args[i] == "-i") && i + 1 < args.Length)
                {
                    inputPath = args[i + 1];
                    i++; // Skip next argument
                }
                else if ((args[i] == "--output" || args[i] == "-o") && i + 1 < args.Length)
                {
                    outputPath = args[i + 1];
                    i++; // Skip next argument
                }
                else if ((args[i] == "--template" || args[i] == "-t") && i + 1 < args.Length)
                {
                    templatePath = args[i + 1];
                    i++; // Skip next argument
                }
                else if (!args[i].StartsWith("-"))
                {
                    // Positional arguments: first is input, second is output, third is template (for create mode)
                    if (inputPath == null)
                        inputPath = args[i];
                    else if (outputPath == null)
                        outputPath = args[i];
                    else if (templatePath == null)
                        templatePath = args[i];
                }
            }

            // Validate arguments
            if (string.IsNullOrEmpty(inputPath) || string.IsNullOrEmpty(outputPath))
            {
                Console.WriteLine("❌ Error: Both input and output paths are required");
                Console.WriteLine();
                ShowCommandLineHelp();
                return 1;
            }

            // Determine mode: if not specified, infer from file extensions
            if (string.IsNullOrEmpty(mode))
            {
                string inputExt = Path.GetExtension(inputPath).ToLower();
                string outputExt = Path.GetExtension(outputPath).ToLower();
                
                if (inputExt == ".pptx" && outputExt == ".json")
                {
                    mode = "extract";
                }
                else if (inputExt == ".json" && outputExt == ".pptx")
                {
                    mode = "create";
                }
                else
                {
                    // Default to create (JSON to PPTX) for backward compatibility
                    mode = "create";
                }
            }

            // Execute based on mode
            if (mode == "extract")
            {
                return ProcessPowerPoint(inputPath, outputPath);
            }
            else if (mode == "create")
            {
                // JSON to PPTX conversion requires template file
                if (string.IsNullOrEmpty(templatePath))
                {
                    Console.WriteLine("❌ Error: Template file path is required for JSON to PPTX conversion.");
                    Console.WriteLine("Usage: OfficeHelperOpenXml.exe --mode create --input <json_file> --output <output_pptx> --template <template_pptx>");
                    Console.WriteLine("Or: OfficeHelperOpenXml.exe <json_file> <output_pptx> <template_pptx>");
                    return 1;
                }
                
                return CreatePPTFromJson(inputPath, templatePath, outputPath);
            }
            else
            {
                Console.WriteLine($"❌ Error: Invalid mode '{mode}'. Use 'extract' or 'create'");
                Console.WriteLine();
                ShowCommandLineHelp();
                return 1;
            }
        }

        /// <summary>
        /// Parse compare command arguments (DISABLED - comparison feature removed)
        /// </summary>
        /// <param name="args">Command-line arguments</param>
        /// <returns>Exit code</returns>
        /*
        private static int ParseCompareCommand(string[] args)
        {
            string generatedPath = null;
            string repairedPath = null;
            string reportPath = null;
            string actionPlanPath = null;

            // Parse arguments
            for (int i = 1; i < args.Length; i++)
            {
                if ((args[i] == "--generated" || args[i] == "-g") && i + 1 < args.Length)
                {
                    generatedPath = args[i + 1];
                    i++;
                }
                else if ((args[i] == "--repaired" || args[i] == "-r") && i + 1 < args.Length)
                {
                    repairedPath = args[i + 1];
                    i++;
                }
                else if ((args[i] == "--report" || args[i] == "-o") && i + 1 < args.Length)
                {
                    reportPath = args[i + 1];
                    i++;
                }
                else if ((args[i] == "--action-plan" || args[i] == "-a") && i + 1 < args.Length)
                {
                    actionPlanPath = args[i + 1];
                    i++;
                }
                else if (!args[i].StartsWith("-"))
                {
                    // Positional arguments
                    if (generatedPath == null)
                        generatedPath = args[i];
                    else if (repairedPath == null)
                        repairedPath = args[i];
                    else if (reportPath == null)
                        reportPath = args[i];
                    else if (actionPlanPath == null)
                        actionPlanPath = args[i];
                }
            }

            // Validate required arguments
            if (string.IsNullOrEmpty(generatedPath) || string.IsNullOrEmpty(repairedPath))
            {
                Console.WriteLine("❌ Error: Both generated and repaired PPTX paths are required");
                Console.WriteLine();
                ShowCompareHelp();
                return 1;
            }

            // Set default output paths if not specified
            if (string.IsNullOrEmpty(reportPath))
            {
                reportPath = "comparison_report.md";
            }

            if (string.IsNullOrEmpty(actionPlanPath))
            {
                actionPlanPath = "action_plan.md";
            }

            // Execute comparison
            return ComparePptxFiles(generatedPath, repairedPath, reportPath, actionPlanPath);
        }
        */

        /// <summary>
        /// Displays command-line help information
        /// </summary>
        private static void ShowCommandLineHelp()
        {
            Console.WriteLine("========================================");
            Console.WriteLine("  OfficeHelperOpenXml - Command Line");
            Console.WriteLine("========================================");
            Console.WriteLine();
            Console.WriteLine("Commands:");
            Console.WriteLine("  extract   Extract PPTX to JSON");
            Console.WriteLine("  create    Create PPTX from JSON (requires template)");
            Console.WriteLine();
            Console.WriteLine("Usage:");
            Console.WriteLine("  Extract (PPTX to JSON):");
            Console.WriteLine("    OfficeHelperOpenXml.exe --mode extract --input <pptx_file> --output <json_file>");
            Console.WriteLine("    OfficeHelperOpenXml.exe -m extract -i <pptx_file> -o <json_file>");
            Console.WriteLine("  Create (JSON to PPTX):");
            Console.WriteLine("    OfficeHelperOpenXml.exe --mode create --input <json_file> --output <pptx_file> --template <template_pptx>");
            Console.WriteLine("    OfficeHelperOpenXml.exe -m create -i <json_file> -o <pptx_file> -t <template_pptx>");
            Console.WriteLine("  Auto-detect mode (by file extension):");
            Console.WriteLine("    OfficeHelperOpenXml.exe --input <pptx_file> --output <json_file>");
            Console.WriteLine("    OfficeHelperOpenXml.exe <pptx_file> <json_file>");
            Console.WriteLine("    OfficeHelperOpenXml.exe <json_file> <output_pptx> <template_pptx>");
            Console.WriteLine();
            Console.WriteLine("Options:");
            Console.WriteLine("  --mode, -m        Operation mode: 'extract' (PPTX->JSON) or 'create' (JSON->PPTX)");
            Console.WriteLine("                    If not specified, mode is auto-detected from file extensions");
            Console.WriteLine("  --input, -i       Path to the input file (PPTX for extract, JSON for create)");
            Console.WriteLine("  --output, -o      Path to the output file (JSON for extract, PPTX for create)");
            Console.WriteLine("  --template, -t    Path to the template PPTX file (required for create mode)");
            Console.WriteLine("  --help, -h        Display this help message");
            Console.WriteLine();
            Console.WriteLine("Examples:");
            Console.WriteLine("  OfficeHelperOpenXml.exe presentation.pptx output.json");
            Console.WriteLine("  OfficeHelperOpenXml.exe -m extract -i input.pptx -o output.json");
            Console.WriteLine("  OfficeHelperOpenXml.exe -m create -i data.json -o output.pptx -t template.pptx");
            Console.WriteLine("  OfficeHelperOpenXml.exe data.json output.pptx template.pptx");
            Console.WriteLine();
        }

        /// <summary>
        /// Displays help for compare command (DISABLED - comparison feature removed)
        /// </summary>
        /*
        private static void ShowCompareHelp()
        {
            Console.WriteLine("========================================");
            Console.WriteLine("  PPTX Comparison Tool - Command Line");
            Console.WriteLine("========================================");
            Console.WriteLine();
            Console.WriteLine("Usage:");
            Console.WriteLine("  OfficeHelperOpenXml.exe compare --generated <file1> --repaired <file2>");
            Console.WriteLine("  OfficeHelperOpenXml.exe compare -g <file1> -r <file2> -o <report> -a <action_plan>");
            Console.WriteLine("  OfficeHelperOpenXml.exe compare <generated> <repaired> [report] [action_plan]");
            Console.WriteLine();
            Console.WriteLine("Options:");
            Console.WriteLine("  --generated, -g   Path to the generated PPTX file (required)");
            Console.WriteLine("  --repaired, -r    Path to the repaired PPTX file (required)");
            Console.WriteLine("  --report, -o      Path to save comparison report (default: comparison_report.md)");
            Console.WriteLine("  --action-plan, -a Path to save action plan (default: action_plan.md)");
            Console.WriteLine();
            Console.WriteLine("Examples:");
            Console.WriteLine("  OfficeHelperOpenXml.exe compare generated.pptx repaired.pptx");
            Console.WriteLine("  OfficeHelperOpenXml.exe compare -g gen.pptx -r fixed.pptx -o report.md -a plan.md");
            Console.WriteLine();
        }
        */

        /// <summary>
        /// 从 JSON 文件创建 PowerPoint 文件
        /// </summary>
        /// <param name="jsonPath">输入 JSON 文件路径</param>
        /// <param name="templatePath">模板 PPTX 文件路径</param>
        /// <param name="outputPath">输出 PPTX 文件路径</param>
        /// <returns>退出代码</returns>
        private static int CreatePPTFromJson(string jsonPath, string templatePath, string outputPath)
        {
            // 验证 JSON 文件是否存在
            if (!File.Exists(jsonPath))
            {
                Console.WriteLine($"❌ 错误: JSON 文件不存在 - {jsonPath}");
                return 1;
            }

            // 验证模板文件是否存在
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"❌ 错误: 模板文件不存在 - {templatePath}");
                return 1;
            }

            // 验证输出目录是否可写
            var outputDirectory = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDirectory) && !Directory.Exists(outputDirectory))
            {
                try
                {
                    Directory.CreateDirectory(outputDirectory);
                    Console.WriteLine($"📁 已创建输出目录: {outputDirectory}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"❌ 错误: 无法创建输出目录 - {ex.Message}");
                    return 1;
                }
            }

            // 检查输出目录是否可写
            try
            {
                var testFile = Path.Combine(outputDirectory ?? ".", "test_write.tmp");
                File.WriteAllText(testFile, "test");
                File.Delete(testFile);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ 错误: 输出目录不可写 - {ex.Message}");
                return 1;
            }

            Console.WriteLine($"📂 开始从 JSON 创建 PowerPoint 文件");
            Console.WriteLine($"📄 输入 JSON: {jsonPath}");
            Console.WriteLine($"📋 模板文件: {templatePath}");
            Console.WriteLine($"💾 输出文件: {outputPath}");
            Console.WriteLine();

            try
            {
                // 读取 JSON 文件
                Console.WriteLine("📖 正在读取 JSON 文件...");
                string jsonData = File.ReadAllText(jsonPath);
                if (string.IsNullOrEmpty(jsonData))
                {
                    Console.WriteLine("❌ 错误: JSON 文件为空");
                    return 1;
                }
                Console.WriteLine($"✅ JSON 文件读取成功 (大小: {jsonData.Length} 字符)");
                Console.WriteLine();

                // 使用 OfficeHelperWrapper 写入 PowerPoint
                Console.WriteLine("🔄 正在处理 PowerPoint 文件...");
                Console.WriteLine("  - 打开模板文件");
                Console.WriteLine("  - 清除现有内容幻灯片");
                Console.WriteLine("  - 写入 JSON 中的 ContentSlides 数据");
                Console.WriteLine("  - 保存到输出路径");
                Console.WriteLine();

                bool success = OfficeHelperWrapper.WritePowerPointFromJson(templatePath, jsonData, outputPath);

                if (success)
                {
                    Console.WriteLine();
                    Console.WriteLine("✅ PowerPoint 文件创建成功！");
                    
                    // 显示文件信息
                    if (File.Exists(outputPath))
                    {
                        FileInfo fileInfo = new FileInfo(outputPath);
                        Console.WriteLine($"📦 输出文件大小: {fileInfo.Length / 1024.0:F2} KB");
                        Console.WriteLine($"📍 输出文件路径: {Path.GetFullPath(outputPath)}");
                    }

                    Console.WriteLine();
                    Console.WriteLine("🎉 处理完成！");
                    return 0;
                }
                else
                {
                    Console.WriteLine();
                    Console.WriteLine("❌ PowerPoint 文件创建失败！");
                    Console.WriteLine("请检查错误日志以获取详细信息。");
                    return 1;
                }
            }
            catch (FileNotFoundException ex)
            {
                Console.WriteLine($"❌ 错误: 文件未找到 - {ex.Message}");
                return 1;
            }
            catch (UnauthorizedAccessException ex)
            {
                Console.WriteLine($"❌ 错误: 访问被拒绝 - {ex.Message}");
                return 1;
            }
            catch (IOException ex)
            {
                Console.WriteLine($"❌ 错误: IO 错误 - {ex.Message}");
                return 1;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ 错误: 处理过程中发生异常");
                Console.WriteLine($"错误信息: {ex.Message}");
                Console.WriteLine($"堆栈跟踪: {ex.StackTrace}");
                return 1;
            }
        }

        /// <summary>
        /// 处理 PowerPoint 文件
        /// </summary>
        /// <param name="pptPath">PowerPoint 文件路径</param>
        /// <param name="outputPath">输出 JSON 文件路径</param>
        /// <returns>退出代码</returns>
        private static int ProcessPowerPoint(string pptPath, string outputPath)
        {
            // 验证文件是否存在
            if (!File.Exists(pptPath))
            {
                Console.WriteLine($"❌ 错误: 文件不存在 - {pptPath}");
                return 1;
            }

            Console.WriteLine($"📂 开始分析 PowerPoint 文件: {pptPath}");
            Console.WriteLine($"📄 输出文件: {outputPath}");
            Console.WriteLine();

            // 使用 OpenXML SDK 进行分析
            using (var reader = PowerPointReaderFactory.CreateReader(pptPath, out bool success))
            {
                if (!success)
                {
                    Console.WriteLine("❌ 加载 PowerPoint 文件失败！");
                    return 1;
                }

                Console.WriteLine("✅ PowerPoint 文件加载成功！");
                Console.WriteLine();

                // 获取分析结果
                Console.WriteLine("📊 正在分析文件内容...");
                var info = reader.PresentationInfo;
                if (info != null)
                {
                    Console.WriteLine($"📑 幻灯片数量: {info.Slides?.Count ?? 0}");
                    Console.WriteLine($"📏 页面尺寸: {info.SlideWidth} x {info.SlideHeight}");
                }

                // 保存到文件
                Console.WriteLine();
                Console.WriteLine("💾 正在保存分析结果...");
                if (reader.SaveToJson(outputPath))
                {
                    Console.WriteLine($"✅ JSON 文件已保存到: {outputPath}");
                    
                    // 显示文件大小
                    FileInfo fileInfo = new FileInfo(outputPath);
                    Console.WriteLine($"📦 文件大小: {fileInfo.Length / 1024.0:F2} KB");
                }
                else
                {
                    Console.WriteLine("❌ 保存 JSON 文件失败！");
                    return 1;
                }

                Console.WriteLine();
                Console.WriteLine("🎉 分析完成！");
                return 0;
            }
        }

        /// <summary>
        /// 处理 Excel 文件
        /// </summary>
        /// <param name="excelPath">Excel 文件路径</param>
        /// <param name="outputPath">输出 JSON 文件路径</param>
        /// <returns>退出代码</returns>
        private static int ProcessExcel(string excelPath, string outputPath)
        {
            // 验证文件是否存在
            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"❌ 错误: 文件不存在 - {excelPath}");
                return 1;
            }

            Console.WriteLine($"📂 开始分析 Excel 文件: {excelPath}");
            Console.WriteLine($"📄 输出文件: {outputPath}");
            Console.WriteLine();

            // 使用 OpenXML SDK 进行分析
            using (var reader = new ExcelReader())
            {
                if (!reader.Load(excelPath))
                {
                    Console.WriteLine("❌ 加载 Excel 文件失败！");
                    return 1;
                }

                Console.WriteLine("✅ Excel 文件加载成功！");
                Console.WriteLine();

                // 获取分析结果
                Console.WriteLine("📊 正在分析文件内容...");
                var sheetNames = reader.GetSheetNames();
                Console.WriteLine($"📋 工作表数量: {sheetNames.Count}");
                Console.WriteLine();

                int totalRows = 0;
                foreach (var sheetName in sheetNames)
                {
                    var data = reader.GetSheetData(sheetName);
                    int rowCount = data.Count;
                    totalRows += rowCount;
                    Console.WriteLine($"  📄 {sheetName}: {rowCount} 行");
                }

                Console.WriteLine();
                Console.WriteLine($"📊 总数据行数: {totalRows}");

                // 保存到文件
                Console.WriteLine();
                Console.WriteLine("💾 正在保存分析结果...");
                var allData = reader.GetAllData();
                var json = JsonConvert.SerializeObject(allData, Formatting.Indented);
                File.WriteAllText(outputPath, json);
                Console.WriteLine($"✅ JSON 文件已保存到: {outputPath}");

                // 显示文件大小
                FileInfo fileInfo = new FileInfo(outputPath);
                Console.WriteLine($"📦 文件大小: {fileInfo.Length / 1024.0:F2} KB");

                Console.WriteLine();
                Console.WriteLine("🎉 Excel 分析完成！");
                return 0;
            }
        }


        /// <summary>
        /// Compare two PPTX files and generate reports (DISABLED - comparison feature removed)
        /// </summary>
        /// <param name="generatedPath">Path to generated PPTX file</param>
        /// <param name="repairedPath">Path to repaired PPTX file</param>
        /// <param name="reportPath">Path to save comparison report</param>
        /// <param name="actionPlanPath">Path to save action plan</param>
        /// <returns>Exit code</returns>
        /*
        private static int ComparePptxFiles(string generatedPath, string repairedPath, string reportPath, string actionPlanPath)
        {
            Console.WriteLine("========================================");
            Console.WriteLine("  PPTX Comparison Tool");
            Console.WriteLine("========================================");
            Console.WriteLine();

            // Validate input files
            if (!File.Exists(generatedPath))
            {
                Console.WriteLine($"❌ Error: Generated PPTX file not found - {generatedPath}");
                return 1;
            }

            if (!File.Exists(repairedPath))
            {
                Console.WriteLine($"❌ Error: Repaired PPTX file not found - {repairedPath}");
                return 1;
            }

            try
            {
                // Create comparison tool
                var comparisonTool = new PptxComparisonTool();

                // Run comparison
                var result = comparisonTool.RunComparison(
                    generatedPath,
                    repairedPath,
                    reportPath,
                    actionPlanPath);

                Console.WriteLine();
                Console.WriteLine("========================================");
                
                if (result.Success)
                {
                    Console.WriteLine("✅ Comparison completed successfully!");
                    Console.WriteLine("========================================");
                    Console.WriteLine();
                    Console.WriteLine("Summary:");
                    Console.WriteLine($"  Total Differences: {result.TotalDifferences}");
                    Console.WriteLine($"  Total Issues: {result.TotalIssues}");
                    Console.WriteLine($"  Generated File Valid: {(result.GeneratedFileValid ? "✓" : "✗")}");
                    Console.WriteLine($"  Repaired File Valid: {(result.RepairedFileValid ? "✓" : "✗")}");
                    Console.WriteLine();
                    Console.WriteLine("Output Files:");
                    Console.WriteLine($"  Report: {result.ReportPath}");
                    Console.WriteLine($"  Action Plan: {result.ActionPlanPath}");
                    
                    return 0;
                }
                else
                {
                    Console.WriteLine("❌ Comparison failed");
                    Console.WriteLine("========================================");
                    Console.WriteLine($"Error: {result.ErrorMessage}");
                    return 1;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine();
                Console.WriteLine("========================================");
                Console.WriteLine("❌ Comparison failed with exception");
                Console.WriteLine("========================================");
                Console.WriteLine($"Error: {ex.Message}");
                Console.WriteLine($"Stack trace: {ex.StackTrace}");
                return 1;
            }
        }
        */

        /// <summary>
        /// 显示使用说明
        /// </summary>
        private static void ShowUsage()
        {
            Console.WriteLine("========================================");
            Console.WriteLine("  OfficeHelperOpenXml - Usage Guide");
            Console.WriteLine("========================================");
            Console.WriteLine();
            Console.WriteLine("This program provides Office file analysis and conversion capabilities.");
            Console.WriteLine();
            Console.WriteLine("Available Features:");
            Console.WriteLine("  1. ProcessPowerPoint  - Analyze PowerPoint files and export to JSON");
            Console.WriteLine("  2. ProcessExcel       - Analyze Excel files and export to JSON");
            Console.WriteLine();
            Console.WriteLine("Note: JSON to PPTX conversion feature has been removed.");
            Console.WriteLine("This project now only supports reading PPTX files and outputting JSON format.");
            Console.WriteLine();
            Console.WriteLine("Command-line Usage:");
            Console.WriteLine("  OfficeHelperOpenXml.exe extract <input_file> <output_file>");
            Console.WriteLine("  OfficeHelperOpenXml.exe --help");
            Console.WriteLine();
            Console.WriteLine("Example:");
            Console.WriteLine("  OfficeHelperOpenXml.exe extract test_ppt\\textbox.pptx output.json");
            Console.WriteLine();
            Console.WriteLine("    * Fill, line, and shadow properties");
            Console.WriteLine("    * Text content with formatting");
            Console.WriteLine("    * Theme colors and color transforms");
            Console.WriteLine();
            Console.WriteLine("========================================");
            Console.WriteLine("Other Features");
            Console.WriteLine("========================================");
            Console.WriteLine();
            Console.WriteLine("To use other features (PowerPoint/Excel analysis):");
            Console.WriteLine("  1. Open Program.cs");
            Console.WriteLine("  2. Uncomment the desired test code in Main method");
            Console.WriteLine("  3. Update file paths to your test files");
            Console.WriteLine("  4. Build and run");
            Console.WriteLine();
            Console.WriteLine("Library Mode:");
            Console.WriteLine("  To use as a library (DLL) instead of executable:");
            Console.WriteLine("  1. Open OfficeHelperOpenXml.csproj");
            Console.WriteLine("  2. Remove or comment out <OutputType>Exe</OutputType>");
            Console.WriteLine("  3. Rebuild to generate DLL");
            Console.WriteLine();
        }
    }
}
