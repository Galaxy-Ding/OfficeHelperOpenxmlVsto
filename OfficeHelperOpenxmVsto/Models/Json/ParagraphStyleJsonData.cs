using Newtonsoft.Json;

namespace OfficeHelperOpenXml.Models.Json
{
    /// <summary>
    /// 表示段落样式数据，用于描述单个段落级别的格式设置
    /// </summary>
    public class ParagraphStyleJsonData
    {
        /// <summary>
        /// 字体大小（以百分之一点为单位，例如 1800 = 18pt）
        /// </summary>
        [JsonProperty("font_size")]
        public int FontSize { get; set; } = 0;

        /// <summary>
        /// 字体名称（Latin字体）
        /// </summary>
        [JsonProperty("font_latin")]
        public string FontLatin { get; set; } = "";

        /// <summary>
        /// 东亚字体
        /// </summary>
        [JsonProperty("font_ea")]
        public string FontEa { get; set; } = "";

        /// <summary>
        /// 复杂脚本字体
        /// </summary>
        [JsonProperty("font_cs")]
        public string FontCs { get; set; } = "";

        /// <summary>
        /// 字体颜色
        /// </summary>
        [JsonProperty("font_color")]
        public string FontColor { get; set; } = "";

        /// <summary>
        /// 主题色名称（如 "tx1", "accent1" 等）
        /// </summary>
        [JsonProperty("scheme_color")]
        public string SchemeColor { get; set; } = "";

        /// <summary>
        /// 颜色变换（用于主题色的亮度调整等）
        /// </summary>
        [JsonProperty("color_transforms")]
        public ColorTransformJsonData ColorTransforms { get; set; }

        /// <summary>
        /// 行间距（以百分比表示，例如 100000 = 100%）
        /// </summary>
        [JsonProperty("line_spacing")]
        public int LineSpacing { get; set; } = 0;

        /// <summary>
        /// 段前间距（以百分比表示）
        /// </summary>
        [JsonProperty("space_before")]
        public int SpaceBefore { get; set; } = 0;

        /// <summary>
        /// 段后间距（以百分比表示）
        /// </summary>
        [JsonProperty("space_after")]
        public int SpaceAfter { get; set; } = 0;

        /// <summary>
        /// 项目符号类型（"none", "bullet", "number" 等）
        /// </summary>
        [JsonProperty("bullet_type")]
        public string BulletType { get; set; } = "none";

        /// <summary>
        /// 项目符号字符（如果是自定义项目符号）
        /// </summary>
        [JsonProperty("bullet_char")]
        public string BulletChar { get; set; } = "";

        /// <summary>
        /// 对齐方式（"left", "center", "right", "justify"）
        /// </summary>
        [JsonProperty("alignment")]
        public string Alignment { get; set; } = "";

        /// <summary>
        /// 缩进级别
        /// </summary>
        [JsonProperty("indent_level")]
        public int IndentLevel { get; set; } = 0;

        /// <summary>
        /// 左缩进（EMU单位）
        /// </summary>
        [JsonProperty("margin_left")]
        public int MarginLeft { get; set; } = 0;

        /// <summary>
        /// 右缩进（EMU单位）
        /// </summary>
        [JsonProperty("margin_right")]
        public int MarginRight { get; set; } = 0;

        /// <summary>
        /// 是否加粗
        /// </summary>
        [JsonProperty("bold")]
        public int Bold { get; set; } = 0;

        /// <summary>
        /// 是否斜体
        /// </summary>
        [JsonProperty("italic")]
        public int Italic { get; set; } = 0;

        /// <summary>
        /// 是否下划线
        /// </summary>
        [JsonProperty("underline")]
        public int Underline { get; set; } = 0;

        /// <summary>
        /// 字符间距（EMU单位）
        /// </summary>
        [JsonProperty("char_spacing")]
        public int CharSpacing { get; set; } = 0;

        /// <summary>
        /// 语言标识（如 "zh-CN"）
        /// </summary>
        [JsonProperty("language")]
        public string Language { get; set; } = "";
    }
}

