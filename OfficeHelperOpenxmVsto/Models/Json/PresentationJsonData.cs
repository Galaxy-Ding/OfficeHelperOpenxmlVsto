using System.Collections.Generic;
using Newtonsoft.Json;

namespace OfficeHelperOpenXml.Models.Json
{
    public class PresentationJsonData
    {
        /// <summary>
        /// 母版幻灯片（保留用于向后兼容，包含母版上的形状）
        /// </summary>
        [JsonProperty("master_slides")]
        public List<SlideJsonData> MasterSlides { get; set; } = new List<SlideJsonData>();

        /// <summary>
        /// 内容幻灯片
        /// </summary>
        [JsonProperty("content_slides")]
        public List<SlideJsonData> ContentSlides { get; set; } = new List<SlideJsonData>();

        /// <summary>
        /// 演示文稿默认文本样式（presentation.xml 中的 defaultTextStyle）
        /// </summary>
        [JsonProperty("default_text_style")]
        public DefaultTextStyleJsonData DefaultTextStyle { get; set; } = new DefaultTextStyleJsonData();

        /// <summary>
        /// 母版样式数据（包含完整的样式定义）
        /// </summary>
        [JsonProperty("slide_master_styles")]
        public List<SlideMasterStyleJsonData> SlideMasterStyles { get; set; } = new List<SlideMasterStyleJsonData>();

        /// <summary>
        /// 幻灯片宽度（厘米）
        /// </summary>
        [JsonProperty("slide_width")]
        public float SlideWidth { get; set; } = 25.4f;

        /// <summary>
        /// 幻灯片高度（厘米）
        /// </summary>
        [JsonProperty("slide_height")]
        public float SlideHeight { get; set; } = 19.05f;
    }
}
