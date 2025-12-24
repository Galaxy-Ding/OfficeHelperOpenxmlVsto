using System;
using System.IO;
using Newtonsoft.Json;
using Xunit;
using OfficeHelperOpenXml.Models.Json;
using OfficeHelperOpenXml.Api;

namespace OfficeHelperOpenXml.Test
{
    public class JsonModelDeserializationTest
    {
        [Fact]
        public void TestDeserializePresentationJsonData()
        {
            // Arrange
            // 使用 TestPaths 获取正确的路径
            var testPptDir = TestPaths.TestPptDir;
            var jsonPath = Path.Combine(testPptDir, "textbox.json");
            
            // 确保 test_ppt 目录存在
            if (!Directory.Exists(testPptDir))
            {
                Directory.CreateDirectory(testPptDir);
            }
            
            // 如果 JSON 文件不存在，先使用 PowerPointReader 生成它
            if (!File.Exists(jsonPath))
            {
                var pptPath = TestPaths.TextboxPptxPath;
                if (!File.Exists(pptPath))
                {
                    throw new FileNotFoundException($"无法找到 textbox.pptx 文件: {pptPath}。请确保文件存在以生成 JSON。");
                }
                
                // 使用 OfficeHelperOpenXmlVsto 生成 JSON 文件
                using (var reader = new PowerPointReader())
                {
                    reader.Load(pptPath);
                    var json = reader.ToJson();
                    File.WriteAllText(jsonPath, json);
                }
            }
            
            // Act
            string jsonContent = File.ReadAllText(jsonPath);
            var presentationData = JsonConvert.DeserializeObject<PresentationJsonData>(jsonContent);
            
            // Assert
            Assert.NotNull(presentationData);
            Assert.NotNull(presentationData.MasterSlides);
            Assert.NotNull(presentationData.ContentSlides);
            Assert.True(presentationData.MasterSlides.Count > 0, "Should have at least one master slide");
            Assert.True(presentationData.ContentSlides.Count > 0, "Should have at least one content slide");
            
            Console.WriteLine($"Successfully deserialized {presentationData.MasterSlides.Count} master slides and {presentationData.ContentSlides.Count} content slides");
        }

        [Fact]
        public void TestShapeBoxParsing()
        {
            // Arrange
            var shape = new ShapeJsonData
            {
                Box = "1.18,3.65,2.60,1.03"
            };
            
            // Act
            bool success = shape.TryParseBox(out float left, out float top, out float width, out float height);
            
            // Assert
            Assert.True(success, "Box parsing should succeed");
            Assert.Equal(1.18f, left, 2);
            Assert.Equal(3.65f, top, 2);
            Assert.Equal(2.60f, width, 2);
            Assert.Equal(1.03f, height, 2);
            
            Console.WriteLine($"Successfully parsed box: left={left}, top={top}, width={width}, height={height}");
        }

        [Fact]
        public void TestShapeBoxParsingInvalidFormat()
        {
            // Arrange
            var shape = new ShapeJsonData
            {
                Box = "invalid"
            };
            
            // Act
            bool success = shape.TryParseBox(out float left, out float top, out float width, out float height);
            
            // Assert
            Assert.False(success, "Box parsing should fail for invalid format");
            Assert.Equal(0, left);
            Assert.Equal(0, top);
            Assert.Equal(0, width);
            Assert.Equal(0, height);
        }

        [Fact]
        public void TestTextRunJsonDataDeserialization()
        {
            // Arrange
            string json = @"{
                ""content"": ""Test Text"",
                ""font"": ""Arial"",
                ""font_size"": 12.0,
                ""font_color"": ""RGB(255, 0, 0)"",
                ""font_bold"": 1,
                ""font_italic"": 0,
                ""font_underline"": 1,
                ""font_strikethrough"": 0
            }";
            
            // Act
            var textRun = JsonConvert.DeserializeObject<TextRunJsonData>(json);
            
            // Assert
            Assert.NotNull(textRun);
            Assert.Equal("Test Text", textRun.Content);
            Assert.Equal("Arial", textRun.Font);
            Assert.Equal(12.0f, textRun.FontSize);
            Assert.Equal("RGB(255, 0, 0)", textRun.FontColor);
            Assert.Equal(1, textRun.FontBold);
            Assert.Equal(0, textRun.FontItalic);
            Assert.Equal(1, textRun.FontUnderline);
            Assert.Equal(0, textRun.FontStrikethrough);
        }

        [Fact]
        public void TestColorTransformDeserialization()
        {
            // Arrange
            string json = @"{
                ""color"": ""RGB(0, 0, 0)"",
                ""opacity"": 1.0,
                ""schemeColor"": ""accent1"",
                ""colorTransforms"": {
                    ""lumMod"": 60000,
                    ""lumOff"": 40000
                }
            }";
            
            // Act
            var fill = JsonConvert.DeserializeObject<FillJsonData>(json);
            
            // Assert
            Assert.NotNull(fill);
            Assert.Equal("accent1", fill.SchemeColor);
            Assert.NotNull(fill.ColorTransforms);
            Assert.Equal(60000, fill.ColorTransforms.LumMod);
            Assert.Equal(40000, fill.ColorTransforms.LumOff);
        }
    }
}
