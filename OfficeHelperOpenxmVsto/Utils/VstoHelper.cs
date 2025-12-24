using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;

namespace OfficeHelperOpenXml.Utils
{
    /// <summary>
    /// VSTO 辅助工具类
    /// </summary>
    public static class VstoHelper
    {
        /// <summary>
        /// 检查 PowerPoint 是否可用
        /// </summary>
        public static bool IsPowerPointAvailable()
        {
            try
            {
                var app = new Application();
                app.Quit();
                Marshal.ReleaseComObject(app);
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 厘米转点（Points）- VSTO 使用点作为单位
        /// </summary>
        public static float CmToPoints(float cm)
        {
            return (float)UnitConverter.CmToPoints(cm);
        }

        /// <summary>
        /// 点转厘米
        /// </summary>
        public static float PointsToCm(float points)
        {
            return (float)UnitConverter.PointsToCm(points);
        }

        /// <summary>
        /// 解析 RGB 颜色字符串为 VSTO 颜色值
        /// 格式：RGB(255, 128, 64)
        /// VSTO 颜色格式：RGB(r, g, b) = r + (g * 256) + (b * 65536)
        /// </summary>
        public static int ParseRgbColor(string rgbString)
        {
            if (string.IsNullOrEmpty(rgbString) || !rgbString.StartsWith("RGB(", StringComparison.OrdinalIgnoreCase))
                return 0; // 黑色

            try
            {
                var content = rgbString.Substring(4, rgbString.Length - 5);
                var parts = content.Split(',');

                if (parts.Length != 3) return 0;

                int r = int.Parse(parts[0].Trim());
                int g = int.Parse(parts[1].Trim());
                int b = int.Parse(parts[2].Trim());

                // 确保值在有效范围内
                r = Math.Max(0, Math.Min(255, r));
                g = Math.Max(0, Math.Min(255, g));
                b = Math.Max(0, Math.Min(255, b));

                // VSTO 颜色格式：RGB(r, g, b) = r + (g * 256) + (b * 65536)
                return r + (g * 256) + (b * 65536);
            }
            catch
            {
                return 0;
            }
        }

        /// <summary>
        /// 安全释放 COM 对象
        /// </summary>
        public static void ReleaseComObject(object obj)
        {
            if (obj != null)
            {
                try
                {
                    Marshal.ReleaseComObject(obj);
                }
                catch
                {
                    // 忽略释放错误
                }
            }
        }

        /// <summary>
        /// 强制垃圾回收以释放 COM 对象
        /// </summary>
        public static void ForceGarbageCollection()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }
    }
}

