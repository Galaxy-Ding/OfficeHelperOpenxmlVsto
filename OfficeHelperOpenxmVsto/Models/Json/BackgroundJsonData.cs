using Newtonsoft.Json;

namespace OfficeHelperOpenXml.Models.Json
{
    /// <summary>
    /// 表示背景样式数据
    /// </summary>
    public class BackgroundJsonData
    {
        /// <summary>
        /// 背景类型（"none", "solid", "gradient", "pattern", "picture"）
        /// </summary>
        [JsonProperty("type")]
        public string Type { get; set; } = "none";

        /// <summary>
        /// 背景填充颜色（用于纯色背景）
        /// </summary>
        [JsonProperty("color")]
        public string Color { get; set; } = "";

        /// <summary>
        /// 主题色名称（如 "bg1", "accent1" 等）
        /// </summary>
        [JsonProperty("scheme_color")]
        public string SchemeColor { get; set; } = "";

        /// <summary>
        /// 颜色变换
        /// </summary>
        [JsonProperty("color_transforms")]
        public ColorTransformJsonData ColorTransforms { get; set; }

        /// <summary>
        /// 渐变填充数据（如果是渐变背景）
        /// </summary>
        [JsonProperty("gradient")]
        public string GradientData { get; set; } = "";

        /// <summary>
        /// 图片数据（Base64，如果是图片背景）
        /// </summary>
        [JsonProperty("picture_base64")]
        public string PictureBase64 { get; set; } = "";

        /// <summary>
        /// 图片格式（"png", "jpg" 等）
        /// </summary>
        [JsonProperty("picture_format")]
        public string PictureFormat { get; set; } = "";

        /// <summary>
        /// 是否应用背景（0=否，1=是）
        /// </summary>
        [JsonProperty("has_background")]
        public int HasBackground { get; set; } = 0;
    }
}

