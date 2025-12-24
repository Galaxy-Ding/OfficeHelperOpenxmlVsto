using System;
using System.IO;

namespace OfficeHelperOpenXml.Test
{
    /// <summary>
    /// 提供测试项目中常用的路径（解决方案根目录、测试输出目录、textbox.pptx 等）
    /// </summary>
    public static class TestPaths
    {
        /// <summary>
        /// 获取解决方案根目录（从测试程序集输出目录向上四级）
        /// </summary>
        public static string GetSolutionRoot()
        {
            var baseDir = AppContext.BaseDirectory;
            return Path.GetFullPath(Path.Combine(baseDir, "..", "..", "..", ".."));
        }

        /// <summary>
        /// 解决方案根目录（与用户工作区 D:\pythonf\c_sharp_project\OfficeHelperOpenxmVsto 对应）
        /// </summary>
        public static string SolutionRoot => GetSolutionRoot();

        /// <summary>
        /// 需要测试的 PowerPoint 文件：textbox.pptx
        /// 路径：D:\pythonf\c_sharp_project\OfficeHelperOpenxmVsto\test_ppt\textbox.pptx
        /// </summary>
        public static string TextboxPptxPath => Path.Combine(SolutionRoot, "test_ppt", "textbox.pptx");

        /// <summary>
        /// 测试输出目录（JSON 等中间结果）
        /// </summary>
        public static string TestOutputDir => Path.Combine(SolutionRoot, "OfficeHelperOpenxmVsto.Test", "test_output");

        /// <summary>
        /// 测试报告目录（Markdown 分析报告）
        /// </summary>
        public static string TestReportsDir => Path.Combine(SolutionRoot, "OfficeHelperOpenxmVsto.Test", "test_reports");

        /// <summary>
        /// textbox.pptx 转换后的 JSON 输出路径
        /// </summary>
        public static string TextboxJsonOutputPath => Path.Combine(TestOutputDir, "textbox.json");

        /// <summary>
        /// 26xdemo2.pptx 模板文件路径
        /// 路径：D:\pythonf\c_sharp_project\OfficeHelperOpenxmVsto\26xdemo2.pptx
        /// </summary>
        public static string Template26xdemo2Path => Path.Combine(SolutionRoot, "26xdemo2.pptx");

        /// <summary>
        /// test_ppt 目录路径
        /// </summary>
        public static string TestPptDir => Path.Combine(SolutionRoot, "test_ppt");
    }
}


