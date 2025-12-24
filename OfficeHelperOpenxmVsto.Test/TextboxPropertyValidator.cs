using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeHelperOpenXml.Models;

namespace OfficeHelperOpenXml.Test
{
    /// <summary>
    /// 属性验证结果
    /// </summary>
    public class PropertyValidationResult
    {
        public bool IsValid { get; set; }
        public string PropertyName { get; set; }
        public object ExpectedValue { get; set; }
        public object ActualValue { get; set; }
        public string Message { get; set; }

        public PropertyValidationResult()
        {
            IsValid = true;
            PropertyName = string.Empty;
            Message = string.Empty;
        }
    }

    /// <summary>
    /// 文本框属性验证工具
    /// </summary>
    public class TextboxPropertyValidator
    {
        private List<PropertyValidationResult> _validationResults;
        private const float FloatTolerance = 0.001f;

        public TextboxPropertyValidator()
        {
            _validationResults = new List<PropertyValidationResult>();
        }

        /// <summary>
        /// 验证属性值是否匹配
        /// </summary>
        public PropertyValidationResult ValidatePropertyMatch<T>(string propertyName, T expected, T actual)
        {
            var result = new PropertyValidationResult
            {
                PropertyName = propertyName,
                ExpectedValue = expected,
                ActualValue = actual
            };

            if (expected == null && actual == null)
            {
                result.IsValid = true;
                result.Message = $"{propertyName}: Both are null (OK)";
            }
            else if (expected == null || actual == null)
            {
                result.IsValid = false;
                result.Message = $"{propertyName}: Expected {(expected == null ? "null" : expected.ToString())}, Actual {(actual == null ? "null" : actual.ToString())}";
            }
            else if (typeof(T) == typeof(float))
            {
                var expectedFloat = Convert.ToSingle(expected);
                var actualFloat = Convert.ToSingle(actual);
                result.IsValid = Math.Abs(expectedFloat - actualFloat) < FloatTolerance;
                result.Message = result.IsValid
                    ? $"{propertyName}: {expectedFloat} ≈ {actualFloat} (OK)"
                    : $"{propertyName}: Expected {expectedFloat}, Actual {actualFloat}";
            }
            else if (expected.Equals(actual))
            {
                result.IsValid = true;
                result.Message = $"{propertyName}: {expected} (OK)";
            }
            else
            {
                result.IsValid = false;
                result.Message = $"{propertyName}: Expected {expected}, Actual {actual}";
            }

            _validationResults.Add(result);
            return result;
        }

        /// <summary>
        /// 比较颜色信息
        /// </summary>
        public List<PropertyValidationResult> CompareColorInfo(string prefix, ColorInfo expected, ColorInfo actual)
        {
            var results = new List<PropertyValidationResult>();

            if (expected == null && actual == null)
            {
                results.Add(new PropertyValidationResult
                {
                    IsValid = true,
                    PropertyName = $"{prefix}.Color",
                    Message = $"{prefix}.Color: Both are null (OK)"
                });
                return results;
            }

            if (expected == null || actual == null)
            {
                results.Add(new PropertyValidationResult
                {
                    IsValid = false,
                    PropertyName = $"{prefix}.Color",
                    ExpectedValue = expected,
                    ActualValue = actual,
                    Message = $"{prefix}.Color: Expected {(expected == null ? "null" : "not null")}, Actual {(actual == null ? "null" : "not null")}"
                });
                return results;
            }

            results.Add(ValidatePropertyMatch($"{prefix}.Red", expected.Red, actual.Red));
            results.Add(ValidatePropertyMatch($"{prefix}.Green", expected.Green, actual.Green));
            results.Add(ValidatePropertyMatch($"{prefix}.Blue", expected.Blue, actual.Blue));
            results.Add(ValidatePropertyMatch($"{prefix}.IsTransparent", expected.IsTransparent, actual.IsTransparent));
            results.Add(ValidatePropertyMatch($"{prefix}.IsThemeColor", expected.IsThemeColor, actual.IsThemeColor));
            results.Add(ValidatePropertyMatch($"{prefix}.SchemeColorIndex", expected.SchemeColorIndex, actual.SchemeColorIndex));
            results.Add(ValidatePropertyMatch($"{prefix}.SchemeColorName", expected.SchemeColorName, actual.SchemeColorName));

            // 比较颜色变换
            if (expected.Transforms != null || actual.Transforms != null)
            {
                if (expected.Transforms == null || actual.Transforms == null)
                {
                    results.Add(new PropertyValidationResult
                    {
                        IsValid = false,
                        PropertyName = $"{prefix}.Transforms",
                        Message = $"{prefix}.Transforms: Expected {(expected.Transforms == null ? "null" : "not null")}, Actual {(actual.Transforms == null ? "null" : "not null")}"
                    });
                }
                else
                {
                    results.Add(ValidatePropertyMatch($"{prefix}.Transforms.LumMod", expected.Transforms.LumMod, actual.Transforms.LumMod));
                    results.Add(ValidatePropertyMatch($"{prefix}.Transforms.LumOff", expected.Transforms.LumOff, actual.Transforms.LumOff));
                    results.Add(ValidatePropertyMatch($"{prefix}.Transforms.Tint", expected.Transforms.Tint, actual.Transforms.Tint));
                    results.Add(ValidatePropertyMatch($"{prefix}.Transforms.Shade", expected.Transforms.Shade, actual.Transforms.Shade));
                }
            }

            return results;
        }

        /// <summary>
        /// 比较阴影信息
        /// </summary>
        public List<PropertyValidationResult> CompareShadowInfo(string prefix, ShadowInfo expected, ShadowInfo actual)
        {
            var results = new List<PropertyValidationResult>();

            if (expected == null && actual == null)
            {
                results.Add(new PropertyValidationResult
                {
                    IsValid = true,
                    PropertyName = $"{prefix}.Shadow",
                    Message = $"{prefix}.Shadow: Both are null (OK)"
                });
                return results;
            }

            if (expected == null || actual == null)
            {
                results.Add(new PropertyValidationResult
                {
                    IsValid = false,
                    PropertyName = $"{prefix}.Shadow",
                    ExpectedValue = expected,
                    ActualValue = actual,
                    Message = $"{prefix}.Shadow: Expected {(expected == null ? "null" : "not null")}, Actual {(actual == null ? "null" : "not null")}"
                });
                return results;
            }

            results.Add(ValidatePropertyMatch($"{prefix}.HasShadow", expected.HasShadow, actual.HasShadow));
            results.Add(ValidatePropertyMatch($"{prefix}.Type", expected.Type, actual.Type));
            results.Add(ValidatePropertyMatch($"{prefix}.Blur", expected.Blur, actual.Blur));
            results.Add(ValidatePropertyMatch($"{prefix}.Distance", expected.Distance, actual.Distance));
            results.Add(ValidatePropertyMatch($"{prefix}.Angle", expected.Angle, actual.Angle));
            results.Add(ValidatePropertyMatch($"{prefix}.Transparency", expected.Transparency, actual.Transparency));

            if (expected.Color != null || actual.Color != null)
            {
                results.AddRange(CompareColorInfo($"{prefix}.Shadow", expected.Color, actual.Color));
            }

            return results;
        }

        /// <summary>
        /// 比较填充信息
        /// </summary>
        public List<PropertyValidationResult> CompareFillInfo(string prefix, TextFillInfo expected, TextFillInfo actual)
        {
            var results = new List<PropertyValidationResult>();

            if (expected == null && actual == null)
            {
                results.Add(new PropertyValidationResult
                {
                    IsValid = true,
                    PropertyName = $"{prefix}.Fill",
                    Message = $"{prefix}.Fill: Both are null (OK)"
                });
                return results;
            }

            if (expected == null || actual == null)
            {
                results.Add(new PropertyValidationResult
                {
                    IsValid = false,
                    PropertyName = $"{prefix}.Fill",
                    ExpectedValue = expected,
                    ActualValue = actual,
                    Message = $"{prefix}.Fill: Expected {(expected == null ? "null" : "not null")}, Actual {(actual == null ? "null" : "not null")}"
                });
                return results;
            }

            results.Add(ValidatePropertyMatch($"{prefix}.HasFill", expected.HasFill, actual.HasFill));
            results.Add(ValidatePropertyMatch($"{prefix}.FillType", expected.FillType, actual.FillType));
            results.Add(ValidatePropertyMatch($"{prefix}.Transparency", expected.Transparency, actual.Transparency));

            if (expected.Color != null || actual.Color != null)
            {
                results.AddRange(CompareColorInfo($"{prefix}.Fill", expected.Color, actual.Color));
            }

            // 比较渐变
            if (expected.Gradient != null || actual.Gradient != null)
            {
                if (expected.Gradient == null || actual.Gradient == null)
                {
                    results.Add(new PropertyValidationResult
                    {
                        IsValid = false,
                        PropertyName = $"{prefix}.Gradient",
                        Message = $"{prefix}.Gradient: Expected {(expected.Gradient == null ? "null" : "not null")}, Actual {(actual.Gradient == null ? "null" : "not null")}"
                    });
                }
                else
                {
                    results.Add(ValidatePropertyMatch($"{prefix}.Gradient.GradientType", expected.Gradient.GradientType, actual.Gradient.GradientType));
                    results.Add(ValidatePropertyMatch($"{prefix}.Gradient.Angle", expected.Gradient.Angle, actual.Gradient.Angle));
                    // 可以进一步比较渐变停止点
                }
            }

            // 比较图案
            if (expected.Pattern != null || actual.Pattern != null)
            {
                if (expected.Pattern == null || actual.Pattern == null)
                {
                    results.Add(new PropertyValidationResult
                    {
                        IsValid = false,
                        PropertyName = $"{prefix}.Pattern",
                        Message = $"{prefix}.Pattern: Expected {(expected.Pattern == null ? "null" : "not null")}, Actual {(actual.Pattern == null ? "null" : "not null")}"
                    });
                }
                else
                {
                    results.Add(ValidatePropertyMatch($"{prefix}.Pattern.PatternType", expected.Pattern.PatternType, actual.Pattern.PatternType));
                }
            }

            return results;
        }

        /// <summary>
        /// 比较轮廓信息
        /// </summary>
        public List<PropertyValidationResult> CompareOutlineInfo(string prefix, TextOutlineInfo expected, TextOutlineInfo actual)
        {
            var results = new List<PropertyValidationResult>();

            if (expected == null && actual == null)
            {
                results.Add(new PropertyValidationResult
                {
                    IsValid = true,
                    PropertyName = $"{prefix}.Outline",
                    Message = $"{prefix}.Outline: Both are null (OK)"
                });
                return results;
            }

            if (expected == null || actual == null)
            {
                results.Add(new PropertyValidationResult
                {
                    IsValid = false,
                    PropertyName = $"{prefix}.Outline",
                    ExpectedValue = expected,
                    ActualValue = actual,
                    Message = $"{prefix}.Outline: Expected {(expected == null ? "null" : "not null")}, Actual {(actual == null ? "null" : "not null")}"
                });
                return results;
            }

            results.Add(ValidatePropertyMatch($"{prefix}.HasOutline", expected.HasOutline, actual.HasOutline));
            results.Add(ValidatePropertyMatch($"{prefix}.Width", expected.Width, actual.Width));
            results.Add(ValidatePropertyMatch($"{prefix}.DashStyle", expected.DashStyle, actual.DashStyle));
            results.Add(ValidatePropertyMatch($"{prefix}.Transparency", expected.Transparency, actual.Transparency));

            if (expected.Color != null || actual.Color != null)
            {
                results.AddRange(CompareColorInfo($"{prefix}.Outline", expected.Color, actual.Color));
            }

            return results;
        }

        /// <summary>
        /// 比较效果信息
        /// </summary>
        public List<PropertyValidationResult> CompareEffectsInfo(string prefix, TextEffectsInfo expected, TextEffectsInfo actual)
        {
            var results = new List<PropertyValidationResult>();

            if (expected == null && actual == null)
            {
                results.Add(new PropertyValidationResult
                {
                    IsValid = true,
                    PropertyName = $"{prefix}.Effects",
                    Message = $"{prefix}.Effects: Both are null (OK)"
                });
                return results;
            }

            if (expected == null || actual == null)
            {
                results.Add(new PropertyValidationResult
                {
                    IsValid = false,
                    PropertyName = $"{prefix}.Effects",
                    ExpectedValue = expected,
                    ActualValue = actual,
                    Message = $"{prefix}.Effects: Expected {(expected == null ? "null" : "not null")}, Actual {(actual == null ? "null" : "not null")}"
                });
                return results;
            }

            results.Add(ValidatePropertyMatch($"{prefix}.HasEffects", expected.HasEffects, actual.HasEffects));
            results.Add(ValidatePropertyMatch($"{prefix}.HasShadow", expected.HasShadow, actual.HasShadow));
            results.Add(ValidatePropertyMatch($"{prefix}.HasGlow", expected.HasGlow, actual.HasGlow));
            results.Add(ValidatePropertyMatch($"{prefix}.HasReflection", expected.HasReflection, actual.HasReflection));
            results.Add(ValidatePropertyMatch($"{prefix}.HasSoftEdge", expected.HasSoftEdge, actual.HasSoftEdge));
            results.Add(ValidatePropertyMatch($"{prefix}.SoftEdgeRadius", expected.SoftEdgeRadius, actual.SoftEdgeRadius));

            // 比较阴影
            if (expected.Shadow != null || actual.Shadow != null)
            {
                results.AddRange(CompareShadowInfo($"{prefix}.Effects", expected.Shadow, actual.Shadow));
            }

            // 比较发光
            if (expected.Glow != null || actual.Glow != null)
            {
                if (expected.Glow == null || actual.Glow == null)
                {
                    results.Add(new PropertyValidationResult
                    {
                        IsValid = false,
                        PropertyName = $"{prefix}.Glow",
                        Message = $"{prefix}.Glow: Expected {(expected.Glow == null ? "null" : "not null")}, Actual {(actual.Glow == null ? "null" : "not null")}"
                    });
                }
                else
                {
                    results.Add(ValidatePropertyMatch($"{prefix}.Glow.Radius", expected.Glow.Radius, actual.Glow.Radius));
                    results.Add(ValidatePropertyMatch($"{prefix}.Glow.Transparency", expected.Glow.Transparency, actual.Glow.Transparency));
                    if (expected.Glow.Color != null || actual.Glow.Color != null)
                    {
                        results.AddRange(CompareColorInfo($"{prefix}.Glow", expected.Glow.Color, actual.Glow.Color));
                    }
                }
            }

            // 比较反射
            if (expected.Reflection != null || actual.Reflection != null)
            {
                if (expected.Reflection == null || actual.Reflection == null)
                {
                    results.Add(new PropertyValidationResult
                    {
                        IsValid = false,
                        PropertyName = $"{prefix}.Reflection",
                        Message = $"{prefix}.Reflection: Expected {(expected.Reflection == null ? "null" : "not null")}, Actual {(actual.Reflection == null ? "null" : "not null")}"
                    });
                }
                else
                {
                    results.Add(ValidatePropertyMatch($"{prefix}.Reflection.BlurRadius", expected.Reflection.BlurRadius, actual.Reflection.BlurRadius));
                    results.Add(ValidatePropertyMatch($"{prefix}.Reflection.Distance", expected.Reflection.Distance, actual.Reflection.Distance));
                    results.Add(ValidatePropertyMatch($"{prefix}.Reflection.StartOpacity", expected.Reflection.StartOpacity, actual.Reflection.StartOpacity));
                }
            }

            return results;
        }

        /// <summary>
        /// 生成验证报告
        /// </summary>
        public string GenerateValidationReport()
        {
            var sb = new StringBuilder();
            sb.AppendLine("# 属性验证报告");
            sb.AppendLine();
            sb.AppendLine($"生成时间: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            sb.AppendLine();

            var total = _validationResults.Count;
            var passed = _validationResults.Count(r => r.IsValid);
            var failed = total - passed;

            sb.AppendLine("## 验证摘要");
            sb.AppendLine($"- 总验证项: {total}");
            sb.AppendLine($"- 通过: {passed}");
            sb.AppendLine($"- 失败: {failed}");
            sb.AppendLine($"- 通过率: {(total > 0 ? (double)passed / total * 100 : 0):F2}%");
            sb.AppendLine();

            if (failed > 0)
            {
                sb.AppendLine("## 失败详情");
                foreach (var result in _validationResults)
                {
                    if (!result.IsValid)
                    {
                        sb.AppendLine($"### {result.PropertyName}");
                        sb.AppendLine($"- 预期值: {result.ExpectedValue}");
                        sb.AppendLine($"- 实际值: {result.ActualValue}");
                        sb.AppendLine($"- 消息: {result.Message}");
                        sb.AppendLine();
                    }
                }
            }

            return sb.ToString();
        }

        /// <summary>
        /// 获取所有验证结果
        /// </summary>
        public List<PropertyValidationResult> GetValidationResults()
        {
            return new List<PropertyValidationResult>(_validationResults);
        }

        /// <summary>
        /// 清除验证结果
        /// </summary>
        public void Clear()
        {
            _validationResults.Clear();
        }
    }
}

