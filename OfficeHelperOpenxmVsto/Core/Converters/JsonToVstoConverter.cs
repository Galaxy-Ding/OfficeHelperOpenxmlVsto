using System;
using Newtonsoft.Json;
using OfficeHelperOpenXml.Models.Json;
using OfficeHelperOpenXml.Utils;

namespace OfficeHelperOpenXml.Core.Converters
{
    /// <summary>
    /// JSON 到 VSTO 转换器
    /// </summary>
    public class JsonToVstoConverter
    {
        /// <summary>
        /// 从 JSON 字符串解析为 PresentationJsonData
        /// </summary>
        public PresentationJsonData ParseJson(string json)
        {
            if (string.IsNullOrEmpty(json))
            {
                throw new ArgumentException("JSON 字符串不能为空", nameof(json));
            }

            try
            {
                var data = JsonConvert.DeserializeObject<PresentationJsonData>(json);
                
                if (data == null)
                {
                    throw new InvalidOperationException("JSON 解析结果为空");
                }

                // 验证数据
                if (!ValidateJsonData(data))
                {
                    var logger = new Logger();
                    logger.LogWarning("JSON 数据验证失败，但将继续处理");
                }

                return data;
            }
            catch (JsonException ex)
            {
                throw new InvalidOperationException($"JSON 解析失败: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// 验证 JSON 数据
        /// </summary>
        public bool ValidateJsonData(PresentationJsonData data)
        {
            if (data == null) return false;

            // 基本验证
            // 可以添加更多验证逻辑
            return true;
        }
    }
}

