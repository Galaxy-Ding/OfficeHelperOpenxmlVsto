using System;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using OfficeHelperOpenXml.Models.Json;
using OfficeHelperOpenXml.Utils;
// 使用别名避免 Shape 类型冲突（Microsoft.Office.Core 和 Microsoft.Office.Interop.PowerPoint 都有 Shape 类型）
using PptShape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace OfficeHelperOpenXml.Core.Writers
{
    /// <summary>
    /// VSTO 样式写入器
    /// </summary>
    public class VstoStyleWriter
    {
        /// <summary>
        /// 将 JSON 中的字体名称映射为当前 Office / 系统中实际可用的字体名称。
        /// 这里给出一个示例映射表，你可以根据自己机器上安装的字体做增删改。
        /// </summary>
        /// <param name="jsonFontName">JSON 中的字体名称</param>
        /// <returns>映射后的字体名称；如果没有映射则返回原始名称或全局默认字体</returns>
        private string ResolveFontName(string jsonFontName)
        {
            if (string.IsNullOrWhiteSpace(jsonFontName))
            {
                // JSON 中没有提供字体时，可按需指定一个全局默认字体
                // 如果希望沿用模板中的字体，可以改为直接返回 null / 空字符串
                const string defaultFontWhenMissing = "宋体";
                return defaultFontWhenMissing;
            }

            // 示例：可以改成读取配置文件或其他方式
            // key：JSON 中的字体名
            // value：系统 / PowerPoint 中实际存在的字体名
            var map = new System.Collections.Generic.Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                // 思源黑体系（示例）
                //{ "思源黑体 CN Heavy", "Source Han Sans CN Heavy" },
                //{ "思源黑体 CN Bold",  "Source Han Sans CN Bold"  },

                // 你可以根据实际安装的字体继续补充 / 修改
            };

            if (map.TryGetValue(jsonFontName, out var resolved))
            {
                return resolved;
            }

            // 默认：没有命中映射时，直接返回原始名称
            return jsonFontName;
        }

        /// <summary>
        /// 应用填充样式
        /// </summary>
        public void ApplyFill(PptShape shape, FillJsonData fillData)
        {
            if (shape == null || fillData == null) return;

            try
            {
                var fill = shape.Fill;
                
                // 首先检查 HasFill 字段
                if (fillData.HasFill == 0)
                {
                    // 无填充
                    fill.Visible = MsoTriState.msoFalse;
                    return;
                }
                
                // 有填充的情况下，检查是否有颜色
                if (!string.IsNullOrEmpty(fillData.Color))
                {
                    int colorValue = VstoHelper.ParseRgbColor(fillData.Color);
                    fill.ForeColor.RGB = colorValue;
                    fill.Visible = MsoTriState.msoTrue;
                    
                    // 设置透明度（0-1 转换为 0-100）
                    if (fillData.Opacity > 0)
                    {
                        fill.Transparency = 1.0f - fillData.Opacity; // VSTO 使用透明度，不是不透明度
                    }
                }
                else
                {
                    // HasFill=1 但没有颜色，仍然显示为无填充
                    fill.Visible = MsoTriState.msoFalse;
                }
            }
            catch (Exception ex)
            {
                // 记录错误但不中断流程
                var logger = new Logger();
                logger.LogWarning($"应用填充样式失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 应用线条样式
        /// </summary>
        public void ApplyLine(PptShape shape, LineJsonData lineData)
        {
            if (shape == null || lineData == null) return;

            try
            {
                var line = shape.Line;
                
                if (lineData.HasOutline == 1)
                {
                    // 有轮廓
                    line.Visible = MsoTriState.msoTrue;
                    
                    // 应用颜色
                    if (!string.IsNullOrEmpty(lineData.Color))
                    {
                        int colorValue = VstoHelper.ParseRgbColor(lineData.Color);
                        line.ForeColor.RGB = colorValue;
                    }
                    
                    // 应用宽度（厘米转点）
                    if (lineData.Width > 0)
                    {
                        line.Weight = VstoHelper.CmToPoints(lineData.Width);
                    }
                }
                else
                {
                    // 无轮廓
                    line.Visible = MsoTriState.msoFalse;
                }
            }
            catch (Exception ex)
            {
                var logger = new Logger();
                logger.LogWarning($"应用线条样式失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 应用阴影样式
        /// </summary>
        public void ApplyShadow(PptShape shape, ShadowJsonData shadowData)
        {
            if (shape == null || shadowData == null) return;

            try
            {
                var shadow = shape.Shadow;
                
                if (shadowData.HasShadow == 1)
                {
                    // 有阴影
                    shadow.Type = MsoShadowType.msoShadow21;
                    shadow.Visible = MsoTriState.msoTrue;
                    
                    // 应用颜色
                    if (!string.IsNullOrEmpty(shadowData.Color))
                    {
                        int colorValue = VstoHelper.ParseRgbColor(shadowData.Color);
                        shadow.ForeColor.RGB = colorValue;
                    }
                    
                    // 应用偏移（厘米转点）
                    if (shadowData.OffsetX != 0)
                    {
                        shadow.OffsetX = VstoHelper.CmToPoints(shadowData.OffsetX);
                    }
                    if (shadowData.OffsetY != 0)
                    {
                        shadow.OffsetY = VstoHelper.CmToPoints(shadowData.OffsetY);
                    }
                    
                    // 应用模糊（厘米转点）
                    if (shadowData.Blur > 0)
                    {
                        shadow.Blur = VstoHelper.CmToPoints(shadowData.Blur);
                    }
                    
                    // 应用透明度
                    if (shadowData.Transparency > 0)
                    {
                        shadow.Transparency = shadowData.Transparency;
                    }
                }
                else
                {
                    // 无阴影
                    shadow.Visible = MsoTriState.msoFalse;
                }
            }
            catch (Exception ex)
            {
                var logger = new Logger();
                logger.LogWarning($"应用阴影样式失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 应用文本样式
        /// </summary>
        public void ApplyTextFormat(TextRange textRange, TextRunJsonData textRun)
        {
            if (textRange == null || textRun == null) return;

            try
            {
                // 字体名称：无论 JSON 是否提供，都统一走 ResolveFontName，
                // 以便为空/缺失时也能套用全局默认字体（例如宋体或公司规范字体）。
                var resolvedFontName = ResolveFontName(textRun.Font);
                if (!string.IsNullOrWhiteSpace(resolvedFontName))
                {
                    // 统一设置多种脚本的字体，保证中文（东亚）、西文等都能应用到期望字体
                    // 如果你希望更精细区分中/英文，也可以在这里按内容拆分处理
                    textRange.Font.Name = resolvedFontName;              // 通用

                    // 东亚字体（中文等），是影响“宋体 / 等线”等显示效果的关键
                    textRange.Font.NameFarEast = resolvedFontName;

                    // 西文字体
                    textRange.Font.NameAscii = resolvedFontName;

                    // 其他脚本（视需要保留或删除）
                    textRange.Font.NameOther = resolvedFontName;
                    textRange.Font.NameComplexScript = resolvedFontName;
                }
                
                // 字体大小
                if (textRun.FontSize > 0)
                {
                    textRange.Font.Size = textRun.FontSize;
                }
                
                // 字体颜色
                if (!string.IsNullOrEmpty(textRun.FontColor))
                {
                    int colorValue = VstoHelper.ParseRgbColor(textRun.FontColor);
                    textRange.Font.Color.RGB = colorValue;
                }
                
                // 粗体
                textRange.Font.Bold = textRun.FontBold == 1 
                    ? MsoTriState.msoTrue 
                    : MsoTriState.msoFalse;
                
                // 斜体
                textRange.Font.Italic = textRun.FontItalic == 1 
                    ? MsoTriState.msoTrue 
                    : MsoTriState.msoFalse;
                
                // 下划线
                if (textRun.FontUnderline == 1)
                {
                    textRange.Font.Underline = MsoTriState.msoTrue;
                }
                
                // 删除线（注意：PowerPoint TextRange.Font 可能不支持 Strikethrough，需要检查）
                // if (textRun.FontStrikethrough == 1)
                // {
                //     textRange.Font.Strikethrough = MsoTriState.msoTrue;
                // }
            }
            catch (Exception ex)
            {
                var logger = new Logger();
                logger.LogWarning($"应用文本样式失败: {ex.Message}");
            }
        }
    }
}
