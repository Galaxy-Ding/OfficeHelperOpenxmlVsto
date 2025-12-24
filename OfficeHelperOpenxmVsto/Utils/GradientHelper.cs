using System;
using DocumentFormat.OpenXml.Packaging;
using OfficeHelperOpenXml.Models;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeHelperOpenXml.Utils
{
    /// <summary>
    /// Helper methods for extracting gradient and pattern fill information
    /// shared between text and shape components.
    /// </summary>
    public static class GradientHelper
    {
        /// <summary>
        /// Extract gradient fill information from an OpenXML GradientFill element.
        /// Logic factored out from TextComponent.ExtractGradientInfo.
        /// </summary>
        public static GradientInfo ExtractGradientInfo(A.GradientFill gradFill, SlidePart slidePart)
        {
            var gradInfo = new GradientInfo();

            if (gradFill == null)
            {
                return gradInfo;
            }

            try
            {
                // Extract gradient stops
                var gsLst = gradFill.GradientStopList;
                if (gsLst != null)
                {
                    foreach (var gs in gsLst.Elements<A.GradientStop>())
                    {
                        if (gs.Position != null)
                        {
                            float position = gs.Position.Value / 100000.0f;
                            ColorInfo color = new ColorInfo();

                            // Extract color from gradient stop - check RgbColorModelHex first
                            var rgbColor = gs.GetFirstChild<A.RgbColorModelHex>();
                            if (rgbColor != null && rgbColor.Val != null)
                            {
                                color = ColorHelper.ParseHexColor(rgbColor.Val.Value);
                                color.OriginalHex = rgbColor.Val.Value;
                            }
                            else
                            {
                                // If no RGB color, check for SchemeColor
                                var schemeColor = gs.GetFirstChild<A.SchemeColor>();
                                if (schemeColor != null)
                                {
                                    color = ColorHelper.ResolveSchemeColor(schemeColor, slidePart);
                                }
                                else
                                {
                                    // Try other color types
                                    var hslColor = gs.GetFirstChild<A.HslColor>();
                                    if (hslColor != null)
                                    {
                                        // For now, create a default color
                                        color = new ColorInfo(0, 0, 0, false);
                                    }
                                }
                            }

                            // Only add stop if we have a valid color
                            if (color != null)
                            {
                                gradInfo.Stops.Add(new GradientStop(position, color));
                            }
                        }
                    }
                }

                // Extract gradient type and angle
                var lin = gradFill.GetFirstChild<A.LinearGradientFill>();
                if (lin != null)
                {
                    gradInfo.GradientType = "Linear";
                    if (lin.Angle != null)
                    {
                        gradInfo.Angle = lin.Angle.Value / 60000.0f;
                    }
                    else
                    {
                        // Default angle if not specified
                        gradInfo.Angle = 0.0f;
                    }
                }
                else
                {
                    var path = gradFill.GetFirstChild<A.PathGradientFill>();
                    if (path != null)
                    {
                        gradInfo.GradientType = path.Path?.Value.ToString() ?? "Path";
                        gradInfo.Angle = 0.0f; // Path gradients don't have angles
                    }
                    else
                    {
                        // Default to Linear if no type specified
                        gradInfo.GradientType = "Linear";
                        gradInfo.Angle = 0.0f;
                    }
                }

                // Validate that we have at least one stop
                if (gradInfo.Stops == null || gradInfo.Stops.Count == 0)
                {
                    Console.WriteLine("警告: 渐变填充没有找到停止点");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"提取渐变填充时出错: {ex.Message}");
                Console.WriteLine($"堆栈跟踪: {ex.StackTrace}");
            }

            return gradInfo;
        }
    }
}


