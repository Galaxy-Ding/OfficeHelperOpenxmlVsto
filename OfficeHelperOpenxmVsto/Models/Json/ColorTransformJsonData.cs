using Newtonsoft.Json;

namespace OfficeHelperOpenXml.Models.Json
{
    public class ColorTransformJsonData
    {
        [JsonProperty("lumMod")]
        public int? LumMod { get; set; }

        [JsonProperty("lumOff")]
        public int? LumOff { get; set; }

        [JsonProperty("tint")]
        public int? Tint { get; set; }

        [JsonProperty("shade")]
        public int? Shade { get; set; }
    }
}
