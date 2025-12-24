using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;
using OfficeHelperOpenXml.Interfaces;
using OfficeHelperOpenXml.Models;
using OfficeHelperOpenXml.Utils;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeHelperOpenXml.Components
{
    /// <summary>
    /// 填充组件 - 使用 OpenXML 提取填充信息
    /// </summary>
    public class FillComponent : IElementComponent
    {
        public string ComponentType => "Fill";
        public bool IsEnabled { get; set; } = true;
        
        public FillInfo Fill { get; set; }
        
        public FillComponent()
        {
            Fill = new FillInfo();
        }
        
        public void ExtractFromShape(Shape shape, SlidePart slidePart)
        {
            try
            {
                Fill = new FillInfo();
                
                // 获取形状属性
                var spPr = shape.ShapeProperties;
                if (spPr == null)
                {
                    Fill.HasFill = false;
                    return;
                }
                
                // 检查是否有 NoFill
                var noFill = spPr.GetFirstChild<A.NoFill>();
                if (noFill != null)
                {
                    Fill.HasFill = false;
                    Fill.FillType = FillType.NoFill;
                    return;
                }
                
                // 检查渐变填充
                var gradFill = spPr.GetFirstChild<A.GradientFill>();
                if (gradFill != null)
                {
                    Fill.HasFill = true;
                    Fill.FillType = FillType.Gradient;
                    Fill.Gradient = GradientHelper.ExtractGradientInfo(gradFill, slidePart);
                    // 设置预览颜色为第一个渐变停止点颜色（如果存在）
                    if (Fill.Gradient != null && Fill.Gradient.Stops != null && Fill.Gradient.Stops.Count > 0)
                    {
                        Fill.Color = Fill.Gradient.Stops[0].Color;
                    }
                    return;
                }

                // 检查实心填充
                var solidFill = spPr.GetFirstChild<A.SolidFill>();
                if (solidFill != null)
                {
                    Fill.HasFill = true;
                    Fill.FillType = FillType.Solid;
                    Fill.Color = ExtractColorFromSolidFill(solidFill, slidePart);
                    return;
                }
                
                // 检查图案填充
                var pattFill = spPr.GetFirstChild<A.PatternFill>();
                if (pattFill != null)
                {
                    Fill.HasFill = true;
                    Fill.FillType = FillType.Pattern;
                    // 仅用于预览的主颜色
                    var fgClr = pattFill.ForegroundColor;
                    if (fgClr != null)
                    {
                        Fill.Color = ExtractColorFromColorType(fgClr, slidePart);
                    }
                    // 完整图案信息
                    Fill.Pattern = ExtractPatternInfo(pattFill, slidePart);
                    return;
                }
                
                // 检查图片填充
                var blipFill = spPr.GetFirstChild<A.BlipFill>();
                if (blipFill != null)
                {
                    Fill.HasFill = true;
                    Fill.FillType = FillType.Picture;
                    return;
                }
                
                // 默认无填充
                Fill.HasFill = false;
                Fill.FillType = FillType.NoFill;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"提取填充信息时出错: {ex.Message}");
                Fill.HasFill = false;
                Fill.Color = new ColorInfo(0, 0, 0, true);
            }
        }
        
        private ColorInfo ExtractColorFromSolidFill(A.SolidFill solidFill, SlidePart slidePart)
        {
            // RGB颜色
            var rgbColor = solidFill.RgbColorModelHex;
            if (rgbColor != null && rgbColor.Val != null)
            {
                return ColorHelper.ParseHexColor(rgbColor.Val.Value);
            }
            
            // sRGB颜色
            var srgbColor = solidFill.RgbColorModelPercentage;
            if (srgbColor != null)
            {
                int r = (int)(srgbColor.RedPortion?.Value ?? 0) * 255 / 100000;
                int g = (int)(srgbColor.GreenPortion?.Value ?? 0) * 255 / 100000;
                int b = (int)(srgbColor.BluePortion?.Value ?? 0) * 255 / 100000;
                return new ColorInfo(r, g, b, false);
            }
            
            // 主题颜色
            var schemeColor = solidFill.SchemeColor;
            if (schemeColor != null)
            {
                return ColorHelper.ResolveSchemeColor(schemeColor, slidePart);
            }
            
            return new ColorInfo(0, 0, 0, true);
        }
        
        private ColorInfo ExtractColorFromGradientStop(A.GradientStop stop, SlidePart slidePart)
        {
            var rgbColor = stop.RgbColorModelHex;
            if (rgbColor != null && rgbColor.Val != null)
            {
                return ColorHelper.ParseHexColor(rgbColor.Val.Value);
            }
            
            var schemeColor = stop.SchemeColor;
            if (schemeColor != null)
            {
                return ColorHelper.ResolveSchemeColor(schemeColor, slidePart);
            }
            
            return new ColorInfo(0, 0, 0, true);
        }
        
        private ColorInfo ExtractColorFromColorType(A.ForegroundColor fgClr, SlidePart slidePart)
        {
            var rgbColor = fgClr.RgbColorModelHex;
            if (rgbColor != null && rgbColor.Val != null)
            {
                return ColorHelper.ParseHexColor(rgbColor.Val.Value);
            }
            
            var schemeColor = fgClr.SchemeColor;
            if (schemeColor != null)
            {
                return ColorHelper.ResolveSchemeColor(schemeColor, slidePart);
            }
            
            return new ColorInfo(0, 0, 0, true);
        }

        /// <summary>
        /// 提取图案填充信息，与 TextComponent 中逻辑保持一致
        /// </summary>
        private PatternInfo ExtractPatternInfo(A.PatternFill pattFill, SlidePart slidePart)
        {
            var pattInfo = new PatternInfo();

            try
            {
                // Pattern 类型
                if (pattFill.Preset != null)
                {
                    pattInfo.PatternType = pattFill.Preset.Value.ToString();
                }

                // 前景色
                var fgColor = pattFill.ForegroundColor;
                if (fgColor != null)
                {
                    var rgbColor = fgColor.GetFirstChild<A.RgbColorModelHex>();
                    if (rgbColor != null && rgbColor.Val != null)
                    {
                        pattInfo.ForegroundColor = ColorHelper.ParseHexColor(rgbColor.Val.Value);
                    }

                    var schemeColor = fgColor.GetFirstChild<A.SchemeColor>();
                    if (schemeColor != null)
                    {
                        pattInfo.ForegroundColor = ColorHelper.ResolveSchemeColor(schemeColor, slidePart);
                    }
                }

                // 背景色
                var bgColor = pattFill.BackgroundColor;
                if (bgColor != null)
                {
                    var rgbColor = bgColor.GetFirstChild<A.RgbColorModelHex>();
                    if (rgbColor != null && rgbColor.Val != null)
                    {
                        pattInfo.BackgroundColor = ColorHelper.ParseHexColor(rgbColor.Val.Value);
                    }

                    var schemeColor = bgColor.GetFirstChild<A.SchemeColor>();
                    if (schemeColor != null)
                    {
                        pattInfo.BackgroundColor = ColorHelper.ResolveSchemeColor(schemeColor, slidePart);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"提取图案填充时出错: {ex.Message}");
            }

            return pattInfo;
        }
        
        public string ToJson()
        {
            if (!IsEnabled) return "null";
            
            var jsonParts = new List<string>();

            // 是否有填充
            jsonParts.Add($"\"has_fill\":{(Fill.HasFill ? 1 : 0)}");

            // 填充类型
            string fillTypeStr = "none";
            switch (Fill.FillType)
            {
                case FillType.Solid:
                    fillTypeStr = "solid";
                    break;
                case FillType.Gradient:
                    fillTypeStr = "gradient";
                    break;
                case FillType.Pattern:
                    fillTypeStr = "pattern";
                    break;
                case FillType.Picture:
                    fillTypeStr = "picture";
                    break;
                case FillType.NoFill:
                case FillType.Background:
                default:
                    fillTypeStr = "none";
                    break;
            }
            jsonParts.Add($"\"fill_type\":\"{fillTypeStr}\"");
            
            // 基本颜色信息（用于预览）
            string colorStr = Fill.Color?.ToString() ?? "";
            jsonParts.Add($"\"color\":\"{colorStr}\"");
            
            float opacity = Fill.HasFill ? 1.0f : 0.0f;
            jsonParts.Add($"\"opacity\":{opacity:F1}");
            
            // 原始主题色信息（用于无损写回）
            if (Fill.Color != null && Fill.Color.IsThemeColor && !string.IsNullOrEmpty(Fill.Color.SchemeColorName))
            {
                jsonParts.Add($"\"schemeColor\":\"{Fill.Color.SchemeColorName}\"");
                
                if (Fill.Color.Transforms != null && Fill.Color.Transforms.HasTransforms)
                {
                    var transformParts = new List<string>();
                    if (Fill.Color.Transforms.LumMod.HasValue)
                        transformParts.Add($"\"lumMod\":{Fill.Color.Transforms.LumMod.Value}");
                    if (Fill.Color.Transforms.LumOff.HasValue)
                        transformParts.Add($"\"lumOff\":{Fill.Color.Transforms.LumOff.Value}");
                    if (Fill.Color.Transforms.Tint.HasValue)
                        transformParts.Add($"\"tint\":{Fill.Color.Transforms.Tint.Value}");
                    if (Fill.Color.Transforms.Shade.HasValue)
                        transformParts.Add($"\"shade\":{Fill.Color.Transforms.Shade.Value}");
                    if (Fill.Color.Transforms.SatMod.HasValue)
                        transformParts.Add($"\"satMod\":{Fill.Color.Transforms.SatMod.Value}");
                    if (Fill.Color.Transforms.SatOff.HasValue)
                        transformParts.Add($"\"satOff\":{Fill.Color.Transforms.SatOff.Value}");
                    if (Fill.Color.Transforms.Alpha.HasValue)
                        transformParts.Add($"\"alpha\":{Fill.Color.Transforms.Alpha.Value}");
                    
                    if (transformParts.Count > 0)
                        jsonParts.Add($"\"colorTransforms\":{{{string.Join(",", transformParts)}}}");
                }
            }
            else if (Fill.Color != null && !string.IsNullOrEmpty(Fill.Color.OriginalHex))
            {
                // 保存原始十六进制值
                jsonParts.Add($"\"originalHex\":\"{Fill.Color.OriginalHex}\"");
            }

            // 渐变填充详细信息
            if (Fill.FillType == FillType.Gradient && Fill.Gradient != null && Fill.Gradient.Stops != null && Fill.Gradient.Stops.Count > 0)
            {
                jsonParts.Add($"\"gradient_type\":\"{Fill.Gradient.GradientType}\"");
                jsonParts.Add($"\"angle\":{Fill.Gradient.Angle:F2}");

                var stopParts = new List<string>();
                foreach (var stop in Fill.Gradient.Stops)
                {
                    if (stop?.Color == null) continue;

                    var stopJsonParts = new List<string>();
                    stopJsonParts.Add($"\"position\":{stop.Position:F4}");
                    stopJsonParts.Add($"\"color\":\"{stop.Color}\"");

                    if (stop.Color.IsThemeColor && !string.IsNullOrEmpty(stop.Color.SchemeColorName))
                    {
                        stopJsonParts.Add($"\"schemeColor\":\"{stop.Color.SchemeColorName}\"");
                    }

                    if (stop.Color.Transforms != null && stop.Color.Transforms.HasTransforms)
                    {
                        var transformParts = new List<string>();
                        if (stop.Color.Transforms.LumMod.HasValue)
                            transformParts.Add($"\"lumMod\":{stop.Color.Transforms.LumMod.Value}");
                        if (stop.Color.Transforms.LumOff.HasValue)
                            transformParts.Add($"\"lumOff\":{stop.Color.Transforms.LumOff.Value}");
                        if (stop.Color.Transforms.Tint.HasValue)
                            transformParts.Add($"\"tint\":{stop.Color.Transforms.Tint.Value}");
                        if (stop.Color.Transforms.Shade.HasValue)
                            transformParts.Add($"\"shade\":{stop.Color.Transforms.Shade.Value}");
                        if (stop.Color.Transforms.SatMod.HasValue)
                            transformParts.Add($"\"satMod\":{stop.Color.Transforms.SatMod.Value}");
                        if (stop.Color.Transforms.SatOff.HasValue)
                            transformParts.Add($"\"satOff\":{stop.Color.Transforms.SatOff.Value}");
                        if (stop.Color.Transforms.Alpha.HasValue)
                            transformParts.Add($"\"alpha\":{stop.Color.Transforms.Alpha.Value}");

                        if (transformParts.Count > 0)
                        {
                            stopJsonParts.Add($"\"colorTransforms\":{{{string.Join(",", transformParts)}}}");
                        }
                    }

                    stopParts.Add("{" + string.Join(",", stopJsonParts) + "}");
                }

                if (stopParts.Count > 0)
                {
                    jsonParts.Add($"\"gradient_stops\":[{string.Join(",", stopParts)}]");
                }
            }

            // 图案填充详细信息
            if (Fill.FillType == FillType.Pattern && Fill.Pattern != null)
            {
                jsonParts.Add($"\"pattern_type\":\"{Fill.Pattern.PatternType}\"");
                if (Fill.Pattern.ForegroundColor != null)
                {
                    jsonParts.Add($"\"pattern_foreground_color\":\"{Fill.Pattern.ForegroundColor}\"");
                }
                if (Fill.Pattern.BackgroundColor != null)
                {
                    jsonParts.Add($"\"pattern_background_color\":\"{Fill.Pattern.BackgroundColor}\"");
                }
            }
            
            return string.Join(",", jsonParts);
        }

    }
}
