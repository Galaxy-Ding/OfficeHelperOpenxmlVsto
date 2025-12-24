using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeHelperOpenXml.Models.Json;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeHelperOpenXml.Core.Readers
{
    /// <summary>
    /// 样式读取器，负责从 PPTX 中提取各种样式信息
    /// </summary>
    public class StyleReader
    {
        /// <summary>
        /// 从 PresentationPart 提取默认文本样式
        /// </summary>
        public DefaultTextStyleJsonData ExtractDefaultTextStyle(PresentationPart presentationPart)
        {
            var result = new DefaultTextStyleJsonData();
            
            try
            {
                Console.WriteLine("[DEBUG] Extracting default text style...");
                var defaultTextStyle = presentationPart.Presentation?.DefaultTextStyle;
                if (defaultTextStyle == null)
                {
                    Console.WriteLine("[WARNING] DefaultTextStyle is null!");
                    result.HasDefaultStyle = 0;
                    return result;
                }

                result.HasDefaultStyle = 1;
                result.Levels = new TextStyleLevelsJsonData();

                // 提取9个级别的样式
                ExtractLevelStyle(defaultTextStyle.Level1ParagraphProperties, result.Levels, 1);
                ExtractLevelStyle(defaultTextStyle.Level2ParagraphProperties, result.Levels, 2);
                ExtractLevelStyle(defaultTextStyle.Level3ParagraphProperties, result.Levels, 3);
                ExtractLevelStyle(defaultTextStyle.Level4ParagraphProperties, result.Levels, 4);
                ExtractLevelStyle(defaultTextStyle.Level5ParagraphProperties, result.Levels, 5);
                ExtractLevelStyle(defaultTextStyle.Level6ParagraphProperties, result.Levels, 6);
                ExtractLevelStyle(defaultTextStyle.Level7ParagraphProperties, result.Levels, 7);
                ExtractLevelStyle(defaultTextStyle.Level8ParagraphProperties, result.Levels, 8);
                ExtractLevelStyle(defaultTextStyle.Level9ParagraphProperties, result.Levels, 9);
                
                Console.WriteLine($"[DEBUG] Default text style extracted: {CountNonNullLevels(result.Levels)} levels");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ERROR] 提取默认文本样式时出错: {ex.Message}");
                Console.WriteLine($"[ERROR] StackTrace: {ex.StackTrace}");
                result.HasDefaultStyle = 0;
            }

            return result;
        }

        /// <summary>
        /// 从 SlideMasterPart 提取母版样式
        /// </summary>
        public SlideMasterStyleJsonData ExtractSlideMasterStyle(SlideMasterPart masterPart, uint masterId)
        {
            var result = new SlideMasterStyleJsonData
            {
                MasterId = masterId
            };

            try
            {
                var slideMaster = masterPart.SlideMaster;
                if (slideMaster == null)
                {
                    Console.WriteLine($"[WARNING] SlideMaster is null for master ID {masterId}");
                    return result;
                }

                // 提取母版名称
                result.Name = slideMaster.CommonSlideData?.Name?.Value ?? "Slide Master";
                Console.WriteLine($"[DEBUG] Extracting master style: {result.Name} (ID: {masterId})");

                // 提取 preserve 属性
                if (slideMaster.Preserve != null)
                {
                    result.Preserve = slideMaster.Preserve.Value ? 1 : 0;
                }

                // 提取背景
                result.Background = ExtractBackground(slideMaster);
                Console.WriteLine($"[DEBUG]   Background: HasBackground={result.Background?.HasBackground}");

                // 提取文本样式
                var textStyles = slideMaster.TextStyles;
                if (textStyles != null)
                {
                    Console.WriteLine($"[DEBUG]   TextStyles found, extracting...");
                    result.TitleStyle = ExtractTextStyleLevels(textStyles.TitleStyle);
                    result.BodyStyle = ExtractTextStyleLevels(textStyles.BodyStyle);
                    result.OtherStyle = ExtractTextStyleLevels(textStyles.OtherStyle);
                    
                    Console.WriteLine($"[DEBUG]   TitleStyle levels: {CountNonNullLevels(result.TitleStyle)}");
                    Console.WriteLine($"[DEBUG]   BodyStyle levels: {CountNonNullLevels(result.BodyStyle)}");
                    Console.WriteLine($"[DEBUG]   OtherStyle levels: {CountNonNullLevels(result.OtherStyle)}");
                }
                else
                {
                    Console.WriteLine($"[WARNING]   TextStyles is null!");
                }

                // 提取颜色方案和字体方案（从主题）
                var themePart = masterPart.ThemePart;
                if (themePart?.Theme != null)
                {
                    result.ColorScheme = themePart.Theme.ThemeElements?.ColorScheme?.Name?.Value ?? "";
                    result.FontScheme = themePart.Theme.ThemeElements?.FontScheme?.Name?.Value ?? "";
                }

                // 统计布局数量
                result.LayoutCount = masterPart.SlideLayoutParts.Count();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"提取母版样式时出错: {ex.Message}");
            }

            return result;
        }

        /// <summary>
        /// 提取背景样式
        /// </summary>
        private BackgroundJsonData ExtractBackground(SlideMaster slideMaster)
        {
            var result = new BackgroundJsonData();

            try
            {
                var background = slideMaster.CommonSlideData?.Background;
                if (background == null)
                {
                    result.HasBackground = 0;
                    return result;
                }

                result.HasBackground = 1;

                // 提取背景填充
                var bgProps = background.BackgroundProperties;

                if (bgProps != null)
                {
                    // 从 BackgroundProperties 提取填充
                    var solidFill = bgProps.Elements<A.SolidFill>().FirstOrDefault();
                    if (solidFill != null)
                    {
                        result.Type = "solid";
                        ExtractColorFromFill(solidFill, result);
                    }
                    else if (bgProps.Elements<A.GradientFill>().Any())
                    {
                        result.Type = "gradient";
                        // TODO: 提取渐变详细信息
                    }
                    else if (bgProps.Elements<A.PatternFill>().Any())
                    {
                        result.Type = "pattern";
                    }
                    else if (bgProps.Elements<A.BlipFill>().Any())
                    {
                        result.Type = "picture";
                        // TODO: 提取图片信息
                    }
                    else if (bgProps.Elements<A.NoFill>().Any())
                    {
                        result.Type = "none";
                        result.HasBackground = 0;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"提取背景时出错: {ex.Message}");
                result.HasBackground = 0;
            }

            return result;
        }

        /// <summary>
        /// 从填充中提取颜色信息
        /// </summary>
        private void ExtractColorFromFill(A.SolidFill solidFill, BackgroundJsonData result)
        {
            if (solidFill.RgbColorModelHex != null)
            {
                result.Color = solidFill.RgbColorModelHex.Val?.Value ?? "";
            }
            else if (solidFill.SchemeColor != null)
            {
                result.SchemeColor = solidFill.SchemeColor.Val?.Value.ToString() ?? "";
                
                // 提取颜色变换
                result.ColorTransforms = ExtractColorTransforms(solidFill.SchemeColor);
            }
        }

        /// <summary>
        /// 提取颜色变换
        /// </summary>
        private ColorTransformJsonData ExtractColorTransforms(A.SchemeColor schemeColor)
        {
            var result = new ColorTransformJsonData();

            foreach (var child in schemeColor.ChildElements)
            {
                if (child is A.LuminanceModulation lumMod)
                {
                    result.LumMod = lumMod.Val?.Value;
                }
                else if (child is A.LuminanceOffset lumOff)
                {
                    result.LumOff = lumOff.Val?.Value;
                }
                else if (child is A.Tint tint)
                {
                    result.Tint = tint.Val?.Value;
                }
                else if (child is A.Shade shade)
                {
                    result.Shade = shade.Val?.Value;
                }
            }

            return result;
        }

        /// <summary>
        /// 提取文本样式级别（TitleStyle, BodyStyle, OtherStyle）
        /// </summary>
        private TextStyleLevelsJsonData ExtractTextStyleLevels(OpenXmlCompositeElement styleElement)
        {
            var result = new TextStyleLevelsJsonData();

            if (styleElement == null)
            {
                Console.WriteLine("[WARNING] styleElement is null in ExtractTextStyleLevels");
                return result;
            }

            try
            {
                Console.WriteLine($"[DEBUG] Extracting levels from {styleElement.LocalName}");
                
                // 使用反射或直接访问来提取各级样式
                var level1 = styleElement.Elements<A.Level1ParagraphProperties>().FirstOrDefault();
                var level2 = styleElement.Elements<A.Level2ParagraphProperties>().FirstOrDefault();
                var level3 = styleElement.Elements<A.Level3ParagraphProperties>().FirstOrDefault();
                var level4 = styleElement.Elements<A.Level4ParagraphProperties>().FirstOrDefault();
                var level5 = styleElement.Elements<A.Level5ParagraphProperties>().FirstOrDefault();
                var level6 = styleElement.Elements<A.Level6ParagraphProperties>().FirstOrDefault();
                var level7 = styleElement.Elements<A.Level7ParagraphProperties>().FirstOrDefault();
                var level8 = styleElement.Elements<A.Level8ParagraphProperties>().FirstOrDefault();
                var level9 = styleElement.Elements<A.Level9ParagraphProperties>().FirstOrDefault();

                if (level1 != null)
                {
                    result.Level1 = ExtractParagraphStyle(level1);
                    Console.WriteLine($"[DEBUG] Level1 extracted: FontSize={result.Level1?.FontSize}");
                }
                if (level2 != null)
                {
                    result.Level2 = ExtractParagraphStyle(level2);
                    Console.WriteLine($"[DEBUG] Level2 extracted: FontSize={result.Level2?.FontSize}");
                }
                if (level3 != null)
                {
                    result.Level3 = ExtractParagraphStyle(level3);
                    Console.WriteLine($"[DEBUG] Level3 extracted: FontSize={result.Level3?.FontSize}");
                }
                if (level4 != null) result.Level4 = ExtractParagraphStyle(level4);
                if (level5 != null) result.Level5 = ExtractParagraphStyle(level5);
                if (level6 != null) result.Level6 = ExtractParagraphStyle(level6);
                if (level7 != null) result.Level7 = ExtractParagraphStyle(level7);
                if (level8 != null) result.Level8 = ExtractParagraphStyle(level8);
                if (level9 != null) result.Level9 = ExtractParagraphStyle(level9);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ERROR] 提取文本样式级别时出错: {ex.Message}");
                Console.WriteLine($"[ERROR] StackTrace: {ex.StackTrace}");
            }

            return result;
        }

        /// <summary>
        /// 从 TextParagraphPropertiesType 提取段落样式
        /// </summary>
        private void ExtractLevelStyle(A.TextParagraphPropertiesType paraProps, TextStyleLevelsJsonData levels, int level)
        {
            if (paraProps == null)
                return;

            var style = ExtractParagraphStyle(paraProps);
            levels.SetLevel(level, style);
        }

        /// <summary>
        /// 提取段落样式
        /// </summary>
        private ParagraphStyleJsonData ExtractParagraphStyle(OpenXmlCompositeElement paraProps)
        {
            var result = new ParagraphStyleJsonData();

            if (paraProps == null)
                return result;

            try
            {
                // 将 OpenXmlCompositeElement 转换为实际的类型来访问属性
                var textParaProps = paraProps as A.TextParagraphPropertiesType;
                if (textParaProps != null)
                {
                    // 提取对齐方式
                    if (textParaProps.Alignment != null)
                    {
                        result.Alignment = ConvertAlignmentToXmlValue(textParaProps.Alignment.Value);
                    }

                    // 提取缩进
                    if (textParaProps.LeftMargin != null)
                    {
                        result.MarginLeft = textParaProps.LeftMargin.Value;
                    }

                    if (textParaProps.RightMargin != null)
                    {
                        result.MarginRight = textParaProps.RightMargin.Value;
                    }

                    if (textParaProps.Indent != null)
                    {
                        result.IndentLevel = textParaProps.Indent.Value;
                    }
                }

                // 提取行间距
                var lineSpacing = paraProps.Elements<A.LineSpacing>().FirstOrDefault();
                if (lineSpacing != null)
                {
                    var spcPct = lineSpacing.Elements<A.SpacingPercent>().FirstOrDefault();
                    if (spcPct?.Val != null)
                    {
                        result.LineSpacing = spcPct.Val.Value;
                    }
                }

                // 提取段前间距
                var spaceBefore = paraProps.Elements<A.SpaceBefore>().FirstOrDefault();
                if (spaceBefore != null)
                {
                    var spcPct = spaceBefore.Elements<A.SpacingPercent>().FirstOrDefault();
                    if (spcPct?.Val != null)
                    {
                        result.SpaceBefore = spcPct.Val.Value;
                    }
                }

                // 提取段后间距
                var spaceAfter = paraProps.Elements<A.SpaceAfter>().FirstOrDefault();
                if (spaceAfter != null)
                {
                    var spcPct = spaceAfter.Elements<A.SpacingPercent>().FirstOrDefault();
                    if (spcPct?.Val != null)
                    {
                        result.SpaceAfter = spcPct.Val.Value;
                    }
                }

                // 提取项目符号
                if (paraProps.Elements<A.NoBullet>().Any())
                {
                    result.BulletType = "none";
                }
                else if (paraProps.Elements<A.AutoNumberedBullet>().Any())
                {
                    result.BulletType = "number";
                }
                else if (paraProps.Elements<A.CharacterBullet>().Any())
                {
                    result.BulletType = "bullet";
                    var charBullet = paraProps.Elements<A.CharacterBullet>().FirstOrDefault();
                    if (charBullet?.Char != null)
                    {
                        result.BulletChar = charBullet.Char.Value;
                    }
                }

                // 提取默认运行属性（字体、字号等）
                var defRPr = paraProps.Elements<A.DefaultRunProperties>().FirstOrDefault();
                if (defRPr != null)
                {
                    ExtractRunProperties(defRPr, result);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ERROR] 提取段落样式时出错: {ex.Message}");
                Console.WriteLine($"[ERROR] Element type: {paraProps?.GetType().Name}");
                Console.WriteLine($"[ERROR] StackTrace: {ex.StackTrace}");
            }

            return result;
        }

        /// <summary>
        /// 提取运行属性（字体、字号、颜色等）
        /// </summary>
        private void ExtractRunProperties(A.DefaultRunProperties runProps, ParagraphStyleJsonData result)
        {
            try
            {
                // 字号
                if (runProps.FontSize != null)
                {
                    result.FontSize = runProps.FontSize.Value;
                }

                // 加粗
                if (runProps.Bold != null)
                {
                    result.Bold = runProps.Bold.Value ? 1 : 0;
                }

                // 斜体
                if (runProps.Italic != null)
                {
                    result.Italic = runProps.Italic.Value ? 1 : 0;
                }

                // 下划线
                if (runProps.Underline != null)
                {
                    result.Underline = runProps.Underline.Value != A.TextUnderlineValues.None ? 1 : 0;
                }

                // 字符间距
                if (runProps.Spacing != null)
                {
                    result.CharSpacing = runProps.Spacing.Value;
                }

                // 语言
                if (runProps.Language != null)
                {
                    result.Language = runProps.Language.Value;
                }

                // 字体
                var latinFont = runProps.Elements<A.LatinFont>().FirstOrDefault();
                if (latinFont?.Typeface != null)
                {
                    result.FontLatin = latinFont.Typeface.Value;
                }

                var eastAsianFont = runProps.Elements<A.EastAsianFont>().FirstOrDefault();
                if (eastAsianFont?.Typeface != null)
                {
                    result.FontEa = eastAsianFont.Typeface.Value;
                }

                var complexScriptFont = runProps.Elements<A.ComplexScriptFont>().FirstOrDefault();
                if (complexScriptFont?.Typeface != null)
                {
                    result.FontCs = complexScriptFont.Typeface.Value;
                }

                // 颜色
                var solidFill = runProps.Elements<A.SolidFill>().FirstOrDefault();
                if (solidFill != null)
                {
                    if (solidFill.RgbColorModelHex != null)
                    {
                        result.FontColor = solidFill.RgbColorModelHex.Val?.Value ?? "";
                    }
                    else if (solidFill.SchemeColor != null)
                    {
                        result.SchemeColor = solidFill.SchemeColor.Val?.Value.ToString() ?? "";
                        result.ColorTransforms = ExtractColorTransforms(solidFill.SchemeColor);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"提取运行属性时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 辅助方法：统计非空级别数量
        /// </summary>
        private int CountNonNullLevels(Models.Json.TextStyleLevelsJsonData levels)
        {
            if (levels == null) return 0;

            int count = 0;
            if (levels.Level1 != null) count++;
            if (levels.Level2 != null) count++;
            if (levels.Level3 != null) count++;
            if (levels.Level4 != null) count++;
            if (levels.Level5 != null) count++;
            if (levels.Level6 != null) count++;
            if (levels.Level7 != null) count++;
            if (levels.Level8 != null) count++;
            if (levels.Level9 != null) count++;
            return count;
        }

        /// <summary>
        /// 将 TextAlignmentTypeValues 枚举转换为正确的 XML 属性值
        /// </summary>
        /// <param name="alignment">OpenXml TextAlignmentTypeValues 枚举值</param>
        /// <returns>有效的 XML alignment 属性值 (l, ctr, r, just, dist, justLow, thaiDist)</returns>
        private static string ConvertAlignmentToXmlValue(A.TextAlignmentTypeValues alignment)
        {
            if (alignment == A.TextAlignmentTypeValues.Left)
                return "l";
            if (alignment == A.TextAlignmentTypeValues.Center)
                return "ctr";
            if (alignment == A.TextAlignmentTypeValues.Right)
                return "r";
            if (alignment == A.TextAlignmentTypeValues.Justified)
                return "just";
            if (alignment == A.TextAlignmentTypeValues.Distributed)
                return "dist";
            if (alignment == A.TextAlignmentTypeValues.JustifiedLow)
                return "justLow";
            if (alignment == A.TextAlignmentTypeValues.ThaiDistributed)
                return "thaiDist";

            // 默认左对齐
            return "l";
        }
    }
}

