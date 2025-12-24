using Newtonsoft.Json;

namespace OfficeHelperOpenXml.Models.Json
{
    /// <summary>
    /// 表示多级文本样式数据，包含最多9个级别的段落样式
    /// 用于 TitleStyle, BodyStyle, OtherStyle
    /// </summary>
    public class TextStyleLevelsJsonData
    {
        [JsonProperty("level_1")]
        public ParagraphStyleJsonData Level1 { get; set; }

        [JsonProperty("level_2")]
        public ParagraphStyleJsonData Level2 { get; set; }

        [JsonProperty("level_3")]
        public ParagraphStyleJsonData Level3 { get; set; }

        [JsonProperty("level_4")]
        public ParagraphStyleJsonData Level4 { get; set; }

        [JsonProperty("level_5")]
        public ParagraphStyleJsonData Level5 { get; set; }

        [JsonProperty("level_6")]
        public ParagraphStyleJsonData Level6 { get; set; }

        [JsonProperty("level_7")]
        public ParagraphStyleJsonData Level7 { get; set; }

        [JsonProperty("level_8")]
        public ParagraphStyleJsonData Level8 { get; set; }

        [JsonProperty("level_9")]
        public ParagraphStyleJsonData Level9 { get; set; }

        /// <summary>
        /// 获取指定级别的样式（1-9）
        /// </summary>
        public ParagraphStyleJsonData GetLevel(int level)
        {
            return level switch
            {
                1 => Level1,
                2 => Level2,
                3 => Level3,
                4 => Level4,
                5 => Level5,
                6 => Level6,
                7 => Level7,
                8 => Level8,
                9 => Level9,
                _ => null
            };
        }

        /// <summary>
        /// 设置指定级别的样式（1-9）
        /// </summary>
        public void SetLevel(int level, ParagraphStyleJsonData style)
        {
            switch (level)
            {
                case 1: Level1 = style; break;
                case 2: Level2 = style; break;
                case 3: Level3 = style; break;
                case 4: Level4 = style; break;
                case 5: Level5 = style; break;
                case 6: Level6 = style; break;
                case 7: Level7 = style; break;
                case 8: Level8 = style; break;
                case 9: Level9 = style; break;
            }
        }

        /// <summary>
        /// 检查是否有任何级别的样式数据
        /// </summary>
        public bool HasAnyLevel()
        {
            return Level1 != null || Level2 != null || Level3 != null ||
                   Level4 != null || Level5 != null || Level6 != null ||
                   Level7 != null || Level8 != null || Level9 != null;
        }
    }
}

