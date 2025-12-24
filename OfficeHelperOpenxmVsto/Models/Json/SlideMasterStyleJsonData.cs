using Newtonsoft.Json;

namespace OfficeHelperOpenXml.Models.Json
{
    /// <summary>
    /// 表示 Slide Master 的完整样式数据
    /// 包含背景、标题样式、正文样式、其他样式等
    /// </summary>
    public class SlideMasterStyleJsonData
    {
        /// <summary>
        /// 母版 ID（与 OpenXML 中的 MasterId 对应）
        /// </summary>
        [JsonProperty("master_id")]
        public uint MasterId { get; set; } = 2147483648;

        /// <summary>
        /// 母版名称
        /// </summary>
        [JsonProperty("name")]
        public string Name { get; set; } = "";

        /// <summary>
        /// 是否保留母版（preserve="1"）
        /// </summary>
        [JsonProperty("preserve")]
        public int Preserve { get; set; } = 1;

        /// <summary>
        /// 背景样式
        /// </summary>
        [JsonProperty("background")]
        public BackgroundJsonData Background { get; set; } = new BackgroundJsonData();

        /// <summary>
        /// 标题样式（Title Style）
        /// </summary>
        [JsonProperty("title_style")]
        public TextStyleLevelsJsonData TitleStyle { get; set; } = new TextStyleLevelsJsonData();

        /// <summary>
        /// 正文样式（Body Style），包含9个级别
        /// </summary>
        [JsonProperty("body_style")]
        public TextStyleLevelsJsonData BodyStyle { get; set; } = new TextStyleLevelsJsonData();

        /// <summary>
        /// 其他样式（Other Style）
        /// </summary>
        [JsonProperty("other_style")]
        public TextStyleLevelsJsonData OtherStyle { get; set; } = new TextStyleLevelsJsonData();

        /// <summary>
        /// 颜色方案名称（如 "Office"）
        /// </summary>
        [JsonProperty("color_scheme")]
        public string ColorScheme { get; set; } = "";

        /// <summary>
        /// 字体方案名称（如 "Office"）
        /// </summary>
        [JsonProperty("font_scheme")]
        public string FontScheme { get; set; } = "";

        /// <summary>
        /// 关联的布局数量
        /// </summary>
        [JsonProperty("layout_count")]
        public int LayoutCount { get; set; } = 0;
    }
}

