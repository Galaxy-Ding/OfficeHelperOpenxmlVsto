using Newtonsoft.Json;

namespace OfficeHelperOpenXml.Models.Json
{
    /// <summary>
    /// 表示演示文稿的默认文本样式（presentation.xml 中的 defaultTextStyle）
    /// 包含9个级别的默认段落样式
    /// </summary>
    public class DefaultTextStyleJsonData
    {
        /// <summary>
        /// 多级样式定义（Level1-Level9）
        /// </summary>
        [JsonProperty("levels")]
        public TextStyleLevelsJsonData Levels { get; set; } = new TextStyleLevelsJsonData();

        /// <summary>
        /// 是否有默认文本样式
        /// </summary>
        [JsonProperty("has_default_style")]
        public int HasDefaultStyle { get; set; } = 0;
    }
}

