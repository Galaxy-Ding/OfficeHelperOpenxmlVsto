using System;
using System.Collections.Generic;
using OfficeHelperOpenXml.Models;

namespace OfficeHelperOpenXml.Test.TestData
{
    /// <summary>
    /// 文本框测试数据模型
    /// </summary>
    public class TextboxTestData
    {
        /// <summary>
        /// 文本框ID（用于标识测试中的文本框）
        /// </summary>
        public string TextboxId { get; set; }

        /// <summary>
        /// 预期的文本内容
        /// </summary>
        public string ExpectedContent { get; set; }

        /// <summary>
        /// 预期的文本运行属性
        /// </summary>
        public TextRunInfo ExpectedProperties { get; set; }

        /// <summary>
        /// 额外的属性（用于存储无法直接映射到TextRunInfo的属性）
        /// </summary>
        public Dictionary<string, object> AdditionalProperties { get; set; }

        /// <summary>
        /// 测试描述
        /// </summary>
        public string Description { get; set; }

        public TextboxTestData()
        {
            TextboxId = string.Empty;
            ExpectedContent = string.Empty;
            ExpectedProperties = new TextRunInfo();
            AdditionalProperties = new Dictionary<string, object>();
            Description = string.Empty;
        }

        public TextboxTestData(string id, string content, TextRunInfo properties, string description = "")
        {
            TextboxId = id;
            ExpectedContent = content;
            ExpectedProperties = properties ?? new TextRunInfo();
            AdditionalProperties = new Dictionary<string, object>();
            Description = description;
        }
    }
}

