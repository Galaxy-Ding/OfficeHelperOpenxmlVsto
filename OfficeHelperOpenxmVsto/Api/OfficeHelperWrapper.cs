using System;
using System.IO;
using System.Collections.Generic;
using OfficeHelperOpenXml.Api.PowerPoint;

namespace OfficeHelperOpenXml.Api
{
    /// <summary>
    /// OfficeHelper包装器 - 提供简单的静态方法，便于从Python和C++调用
    /// </summary>
    public static class OfficeHelperWrapper
    {
        private static readonly string Version = "2.0.0-OpenXML";

        /// <summary>
        /// ����PowerPoint�ļ�������JSON�ַ���
        /// </summary>
        /// <param name="filePath">PowerPoint�ļ�·��</param>
        /// <returns>JSON�ַ���</returns>
        public static string AnalyzePowerPoint(string filePath)
        {
            try
            {
                if (string.IsNullOrEmpty(filePath))
                {
                    return "{\"error\":\"�ļ�·������Ϊ��\"}";
                }

                if (!File.Exists(filePath))
                {
                    return $"{{\"error\":\"�ļ�������: {filePath}\"}}";
                }

                using (var reader = new PowerPointReader())
                {
                    if (!reader.Load(filePath))
                    {
                        return "{\"error\":\"�����ļ�ʧ��\"}";
                    }
                    return reader.ToJson();
                }
            }
            catch (Exception ex)
            {
                return $"{{\"error\":\"�����ļ�ʱ����: {ex.Message}\"}}";
            }
        }

        /// <summary>
        /// ����PowerPoint�ļ�������ΪJSON�ļ�
        /// </summary>
        /// <param name="filePath">PowerPoint�ļ�·��</param>
        /// <param name="outputPath">���JSON�ļ�·��</param>
        /// <returns>�Ƿ�ɹ�</returns>
        public static bool AnalyzePowerPointToFile(string filePath, string outputPath)
        {
            try
            {
                if (string.IsNullOrEmpty(filePath) || string.IsNullOrEmpty(outputPath))
                {
                    return false;
                }

                if (!File.Exists(filePath))
                {
                    return false;
                }

                using (var reader = new PowerPointReader())
                {
                    if (!reader.Load(filePath))
                    {
                        return false;
                    }
                    return reader.SaveToJson(outputPath);
                }
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// ����ļ��Ƿ����
        /// </summary>
        /// <param name="filePath">�ļ�·��</param>
        /// <returns>�ļ��Ƿ����</returns>
        public static bool FileExists(string filePath)
        {
            return File.Exists(filePath);
        }

        /// <summary>
        /// ��ȡ�汾��Ϣ
        /// </summary>
        /// <returns>�汾�ַ���</returns>
        public static string GetVersion()
        {
            return Version;
        }

        /// <summary>
        /// 获取库信息
        /// </summary>
        /// <returns>库信息JSON</returns>
        public static string GetLibraryInfo()
        {
            return $"{{\"name\":\"OfficeHelperOpenXml\",\"version\":\"{Version}\",\"framework\":\"OpenXML SDK\",\"targetFramework\":\"netstandard2.0\"}}";
        }

        /// <summary>
        /// 从 JSON 写入 PowerPoint 文件（使用 VSTO）
        /// </summary>
        /// <param name="templatePath">模板文件路径</param>
        /// <param name="jsonData">JSON 数据字符串</param>
        /// <param name="outputPath">输出文件路径</param>
        /// <returns>是否成功</returns>
        public static bool WritePowerPointFromJson(string templatePath, string jsonData, string outputPath)
        {
            try
            {
                if (string.IsNullOrEmpty(templatePath))
                {
                    return false;
                }

                if (string.IsNullOrEmpty(jsonData))
                {
                    return false;
                }

                if (string.IsNullOrEmpty(outputPath))
                {
                    return false;
                }

                if (!File.Exists(templatePath))
                {
                    return false;
                }

                using (var writer = PowerPointWriterFactory.CreateWriter())
                {
                    if (!writer.OpenFromTemplate(templatePath))
                    {
                        return false;
                    }

                    if (!writer.ClearAllContentSlides())
                    {
                        return false;
                    }

                    if (!writer.WriteFromJson(jsonData))
                    {
                        return false;
                    }

                    // 添加日志以便跟踪
                    var logger = new OfficeHelperOpenXml.Utils.Logger();
                    logger.LogInfo("准备保存文件...");
                    
                    bool saveResult = writer.SaveAs(outputPath);
                    
                    if (saveResult)
                    {
                        logger.LogInfo("SaveAs 返回成功");
                    }
                    else
                    {
                        logger.LogError("SaveAs 返回失败");
                    }
                    
                    logger.LogInfo("准备退出 using 块，将自动调用 Dispose() 释放资源");
                    // using 块结束时会自动调用 writer.Dispose()
                    return saveResult;
                }
            }
            catch
            {
                // 记录错误但不抛出异常
                return false;
            }
        }
    }
}
