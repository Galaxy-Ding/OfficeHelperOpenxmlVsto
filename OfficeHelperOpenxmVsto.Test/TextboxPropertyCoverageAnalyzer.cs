using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeHelperOpenXml.Test
{
    /// <summary>
    /// 属性覆盖率检查清单
    /// </summary>
    public class PropertyChecklist
    {
        // 基础属性
        public bool FontName { get; set; }
        public bool FontSize { get; set; }
        public bool FontColor { get; set; }
        public bool IsBold { get; set; }
        public bool IsItalic { get; set; }
        public bool IsUnderline { get; set; }
        public bool IsStrikethrough { get; set; }

        // 阴影属性
        public bool ShadowType { get; set; }
        public bool ShadowColor { get; set; }
        public bool ShadowBlur { get; set; }
        public bool ShadowDistance { get; set; }
        public bool ShadowAngle { get; set; }
        public bool ShadowTransparency { get; set; }

        // 填充属性
        public bool FillType { get; set; }
        public bool FillColor { get; set; }
        public bool FillGradient { get; set; }
        public bool FillPattern { get; set; }
        public bool FillTransparency { get; set; }

        // 轮廓属性
        public bool OutlineWidth { get; set; }
        public bool OutlineColor { get; set; }
        public bool OutlineDashStyle { get; set; }
        public bool OutlineTransparency { get; set; }

        // 效果属性
        public bool Glow { get; set; }
        public bool Reflection { get; set; }
        public bool SoftEdge { get; set; }

        // 其他属性
        public bool HighlightColor { get; set; }
        public bool CharacterSpacing { get; set; }
        public bool Superscript { get; set; }
        public bool Subscript { get; set; }
        public bool ThemeColor { get; set; }
        public bool ColorTransforms { get; set; }

        /// <summary>
        /// 获取已测试的属性数量
        /// </summary>
        public int GetTestedCount()
        {
            return GetType().GetProperties()
                .Where(p => p.PropertyType == typeof(bool))
                .Count(p => (bool)p.GetValue(this));
        }

        /// <summary>
        /// 获取总属性数量
        /// </summary>
        public int GetTotalCount()
        {
            return GetType().GetProperties()
                .Count(p => p.PropertyType == typeof(bool));
        }

        /// <summary>
        /// 获取覆盖率百分比
        /// </summary>
        public double GetCoveragePercentage()
        {
            var total = GetTotalCount();
            if (total == 0) return 0;
            return (double)GetTestedCount() / total * 100;
        }
    }

    /// <summary>
    /// 文本框属性覆盖率分析工具
    /// </summary>
    public class TextboxPropertyCoverageAnalyzer
    {
        private PropertyChecklist _checklist;

        public TextboxPropertyCoverageAnalyzer()
        {
            _checklist = new PropertyChecklist();
        }

        /// <summary>
        /// 分析属性覆盖率
        /// </summary>
        public PropertyChecklist AnalyzePropertyCoverage(Dictionary<string, bool> testedProperties)
        {
            var checklist = new PropertyChecklist();

            // 辅助方法：获取值或默认值
            bool GetValue(string key, bool defaultValue)
            {
                bool value;
                return testedProperties.TryGetValue(key, out value) ? value : defaultValue;
            }

            // 基础属性
            checklist.FontName = GetValue("FontName", false);
            checklist.FontSize = GetValue("FontSize", false);
            checklist.FontColor = GetValue("FontColor", false);
            checklist.IsBold = GetValue("IsBold", false);
            checklist.IsItalic = GetValue("IsItalic", false);
            checklist.IsUnderline = GetValue("IsUnderline", false);
            checklist.IsStrikethrough = GetValue("IsStrikethrough", false);

            // 阴影属性
            checklist.ShadowType = GetValue("ShadowType", false);
            checklist.ShadowColor = GetValue("ShadowColor", false);
            checklist.ShadowBlur = GetValue("ShadowBlur", false);
            checklist.ShadowDistance = GetValue("ShadowDistance", false);
            checklist.ShadowAngle = GetValue("ShadowAngle", false);
            checklist.ShadowTransparency = GetValue("ShadowTransparency", false);

            // 填充属性
            checklist.FillType = GetValue("FillType", false);
            checklist.FillColor = GetValue("FillColor", false);
            checklist.FillGradient = GetValue("FillGradient", false);
            checklist.FillPattern = GetValue("FillPattern", false);
            checklist.FillTransparency = GetValue("FillTransparency", false);

            // 轮廓属性
            checklist.OutlineWidth = GetValue("OutlineWidth", false);
            checklist.OutlineColor = GetValue("OutlineColor", false);
            checklist.OutlineDashStyle = GetValue("OutlineDashStyle", false);
            checklist.OutlineTransparency = GetValue("OutlineTransparency", false);

            // 效果属性
            checklist.Glow = GetValue("Glow", false);
            checklist.Reflection = GetValue("Reflection", false);
            checklist.SoftEdge = GetValue("SoftEdge", false);

            // 其他属性
            checklist.HighlightColor = GetValue("HighlightColor", false);
            checklist.CharacterSpacing = GetValue("CharacterSpacing", false);
            checklist.Superscript = GetValue("Superscript", false);
            checklist.Subscript = GetValue("Subscript", false);
            checklist.ThemeColor = GetValue("ThemeColor", false);
            checklist.ColorTransforms = GetValue("ColorTransforms", false);

            _checklist = checklist;
            return checklist;
        }

        /// <summary>
        /// 生成覆盖率报告
        /// </summary>
        public string GenerateCoverageReport()
        {
            var sb = new StringBuilder();
            sb.AppendLine("# 文本框属性覆盖率报告");
            sb.AppendLine();
            sb.AppendLine($"生成时间: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            sb.AppendLine();

            var tested = _checklist.GetTestedCount();
            var total = _checklist.GetTotalCount();
            var percentage = _checklist.GetCoveragePercentage();

            sb.AppendLine($"## 总体统计");
            sb.AppendLine($"- 总属性数: {total}");
            sb.AppendLine($"- 已测试属性数: {tested}");
            sb.AppendLine($"- 未测试属性数: {total - tested}");
            sb.AppendLine($"- 覆盖率: {percentage:F2}%");
            sb.AppendLine();

            sb.AppendLine("## 属性覆盖清单");
            sb.AppendLine();

            // 基础属性
            sb.AppendLine("### 基础属性 (8项)");
            sb.AppendLine($"- [{( _checklist.FontName ? "x" : " ")}] 字体名称 (FontName)");
            sb.AppendLine($"- [{( _checklist.FontSize ? "x" : " ")}] 字体大小 (FontSize)");
            sb.AppendLine($"- [{( _checklist.FontColor ? "x" : " ")}] 字体颜色 (FontColor)");
            sb.AppendLine($"- [{( _checklist.IsBold ? "x" : " ")}] 粗体 (IsBold)");
            sb.AppendLine($"- [{( _checklist.IsItalic ? "x" : " ")}] 斜体 (IsItalic)");
            sb.AppendLine($"- [{( _checklist.IsUnderline ? "x" : " ")}] 下划线 (IsUnderline)");
            sb.AppendLine($"- [{( _checklist.IsStrikethrough ? "x" : " ")}] 删除线 (IsStrikethrough)");
            sb.AppendLine($"- [{( _checklist.HighlightColor ? "x" : " ")}] 高亮颜色 (HighlightColor)");
            sb.AppendLine();

            // 阴影属性
            sb.AppendLine("### 阴影属性 (6项)");
            sb.AppendLine($"- [{( _checklist.ShadowType ? "x" : " ")}] 阴影类型 (ShadowType)");
            sb.AppendLine($"- [{( _checklist.ShadowColor ? "x" : " ")}] 阴影颜色 (ShadowColor)");
            sb.AppendLine($"- [{( _checklist.ShadowBlur ? "x" : " ")}] 模糊半径 (Blur)");
            sb.AppendLine($"- [{( _checklist.ShadowDistance ? "x" : " ")}] 距离 (Distance)");
            sb.AppendLine($"- [{( _checklist.ShadowAngle ? "x" : " ")}] 角度 (Angle)");
            sb.AppendLine($"- [{( _checklist.ShadowTransparency ? "x" : " ")}] 透明度 (Transparency)");
            sb.AppendLine();

            // 填充属性
            sb.AppendLine("### 填充属性 (5项)");
            sb.AppendLine($"- [{( _checklist.FillType ? "x" : " ")}] 填充类型 (FillType)");
            sb.AppendLine($"- [{( _checklist.FillColor ? "x" : " ")}] 填充颜色 (FillColor)");
            sb.AppendLine($"- [{( _checklist.FillGradient ? "x" : " ")}] 渐变填充 (Gradient)");
            sb.AppendLine($"- [{( _checklist.FillPattern ? "x" : " ")}] 图案填充 (Pattern)");
            sb.AppendLine($"- [{( _checklist.FillTransparency ? "x" : " ")}] 填充透明度 (Transparency)");
            sb.AppendLine();

            // 轮廓属性
            sb.AppendLine("### 轮廓属性 (4项)");
            sb.AppendLine($"- [{( _checklist.OutlineWidth ? "x" : " ")}] 轮廓宽度 (Width)");
            sb.AppendLine($"- [{( _checklist.OutlineColor ? "x" : " ")}] 轮廓颜色 (Color)");
            sb.AppendLine($"- [{( _checklist.OutlineDashStyle ? "x" : " ")}] 虚线样式 (DashStyle)");
            sb.AppendLine($"- [{( _checklist.OutlineTransparency ? "x" : " ")}] 轮廓透明度 (Transparency)");
            sb.AppendLine();

            // 效果属性
            sb.AppendLine("### 效果属性 (3项)");
            sb.AppendLine($"- [{( _checklist.Glow ? "x" : " ")}] 发光 (Glow)");
            sb.AppendLine($"- [{( _checklist.Reflection ? "x" : " ")}] 反射 (Reflection)");
            sb.AppendLine($"- [{( _checklist.SoftEdge ? "x" : " ")}] 柔边 (SoftEdge)");
            sb.AppendLine();

            // 其他属性
            sb.AppendLine("### 其他属性 (5项)");
            sb.AppendLine($"- [{( _checklist.HighlightColor ? "x" : " ")}] 高亮颜色 (HighlightColor)");
            sb.AppendLine($"- [{( _checklist.CharacterSpacing ? "x" : " ")}] 字符间距 (CharacterSpacing)");
            sb.AppendLine($"- [{( _checklist.Superscript ? "x" : " ")}] 上标 (Superscript)");
            sb.AppendLine($"- [{( _checklist.Subscript ? "x" : " ")}] 下标 (Subscript)");
            sb.AppendLine($"- [{( _checklist.ThemeColor ? "x" : " ")}] 主题颜色 (ThemeColor)");
            sb.AppendLine($"- [{( _checklist.ColorTransforms ? "x" : " ")}] 颜色变换 (ColorTransforms)");
            sb.AppendLine();

            // 未测试属性列表
            var untested = GetUntestedProperties();
            if (untested.Count > 0)
            {
                sb.AppendLine("## 未测试属性列表");
                foreach (var prop in untested)
                {
                    sb.AppendLine($"- {prop}");
                }
                sb.AppendLine();
            }

            return sb.ToString();
        }

        /// <summary>
        /// 获取未测试的属性列表
        /// </summary>
        public List<string> GetUntestedProperties()
        {
            var untested = new List<string>();
            var properties = _checklist.GetType().GetProperties()
                .Where(p => p.PropertyType == typeof(bool));

            foreach (var prop in properties)
            {
                if (!(bool)prop.GetValue(_checklist))
                {
                    untested.Add(prop.Name);
                }
            }

            return untested;
        }

        /// <summary>
        /// 获取已测试的属性列表
        /// </summary>
        public List<string> GetTestedProperties()
        {
            var tested = new List<string>();
            var properties = _checklist.GetType().GetProperties()
                .Where(p => p.PropertyType == typeof(bool));

            foreach (var prop in properties)
            {
                if ((bool)prop.GetValue(_checklist))
                {
                    tested.Add(prop.Name);
                }
            }

            return tested;
        }
    }
}

