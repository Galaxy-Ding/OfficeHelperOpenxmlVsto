using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using OfficeHelperOpenXml.Api;
using Xunit;
using Xunit.Abstractions;

namespace OfficeHelperOpenXml.Test
{
    /// <summary>
    /// Validates that text run merging works correctly with real PowerPoint files
    /// Task 8: Validate with real PowerPoint files
    /// Requirements: 4.1, 4.3
    /// </summary>
    public class ValidateMergingTest
    {
        private readonly ITestOutputHelper _output;

        public ValidateMergingTest(ITestOutputHelper output)
        {
            _output = output;
        }

        [Fact]
        public void TestTextboxPptx_GenerateAndInspectJson()
        {
            // Arrange
            string testDir = AppDomain.CurrentDomain.BaseDirectory;
            string solutionRoot = Path.GetFullPath(Path.Combine(testDir, "..", "..", "..", ".."));
            string pptPath = Path.Combine(solutionRoot, "test_ppt", "textbox.pptx");
            string outputPath = Path.Combine(solutionRoot, "test_ppt", "textbox_merged_output.json");
            
            if (!File.Exists(pptPath))
            {
                throw new FileNotFoundException($"Test file not found: {pptPath}");
            }
            
            // Act
            string json;
            using (var reader = new PowerPointReader())
            {
                reader.Load(pptPath);
                json = reader.ToJson();
            }
            
            // Save the output for manual inspection
            File.WriteAllText(outputPath, json);
            _output.WriteLine($"✓ JSON output saved to: {outputPath}");
            _output.WriteLine($"✓ JSON length: {json.Length} characters");
            
            // Parse and inspect the JSON structure
            using (JsonDocument doc = JsonDocument.Parse(json))
            {
                var root = doc.RootElement;
                
                // Log the top-level properties
                _output.WriteLine("\nTop-level JSON properties:");
                foreach (var prop in root.EnumerateObject())
                {
                    _output.WriteLine($"  - {prop.Name}");
                }
                
                // Find and count gradient text runs & gradient-filled shapes
                int gradientTextCount = 0;
                int gradientShapeFillCount = 0;
                int totalTextRuns = 0;
                var gradientTextRuns = new List<(string content, string fillType)>();
                var gradientShapeFills = new List<(string shapeName, string fillType)>();
                
                // Check content_slides (not "slides")
                if (root.TryGetProperty("content_slides", out var contentSlides))
                {
                    foreach (var slide in contentSlides.EnumerateArray())
                    {
                        if (slide.TryGetProperty("shapes", out var shapes))
                        {
                            foreach (var shape in shapes.EnumerateArray())
                            {
                                // 先检查形状本身是否是渐变填充
                                if (shape.TryGetProperty("fill", out var shapeFill))
                                {
                                    if (shapeFill.TryGetProperty("fill_type", out var shapeFillType))
                                    {
                                        string shapeFillTypeStr = shapeFillType.GetString();
                                        if (shapeFillTypeStr == "gradient")
                                        {
                                            gradientShapeFillCount++;

                                            string shapeName = "";
                                            if (shape.TryGetProperty("name", out var nameProp))
                                            {
                                                shapeName = nameProp.GetString() ?? "";
                                            }

                                            gradientShapeFills.Add((shapeName, shapeFillTypeStr));
                                            _output.WriteLine($"\nFound gradient shape fill: '{shapeName}'");
                                            _output.WriteLine($"  - Shape fill type: {shapeFillTypeStr}");
                                        }
                                    }
                                }

                                if (shape.TryGetProperty("text", out var textArray))
                                {
                                    foreach (var textRun in textArray.EnumerateArray())
                                    {
                                        totalTextRuns++;
                                        
                                        // Check if this run has gradient fill
                                        if (textRun.TryGetProperty("text_fill", out var textFill))
                                        {
                                            if (textFill.TryGetProperty("fill_type", out var fillType))
                                            {
                                                string fillTypeStr = fillType.GetString();
                                                if (fillTypeStr == "gradient")
                                                {
                                                    gradientTextCount++;
                                                    
                                                    string contentStr = "";
                                                    if (textRun.TryGetProperty("content", out var content))
                                                    {
                                                        contentStr = content.GetString() ?? "";
                                                    }
                                                    
                                                    gradientTextRuns.Add((contentStr, fillTypeStr));
                                                    _output.WriteLine($"\nFound gradient text run: '{contentStr}'");
                                                    _output.WriteLine($"  - Fill type: {fillTypeStr}");
                                                    
                                                    // Verify gradient properties exist
                                                    if (textFill.TryGetProperty("gradient_type", out var gradType))
                                                    {
                                                        _output.WriteLine($"  - Gradient type: {gradType.GetString()}");
                                                    }
                                                    if (textFill.TryGetProperty("stops", out var stops))
                                                    {
                                                        _output.WriteLine($"  - Gradient stops count: {stops.GetArrayLength()}");
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                
                _output.WriteLine($"\n✓ Total text runs: {totalTextRuns}");
                _output.WriteLine($"✓ Gradient text runs found: {gradientTextCount}");
                _output.WriteLine($"✓ Gradient shape fills found: {gradientShapeFillCount}");
                
                // Log all gradient text runs found
                if (gradientTextRuns.Count > 0)
                {
                    _output.WriteLine("\nAll gradient text runs:");
                    foreach (var (content, fillType) in gradientTextRuns)
                    {
                        _output.WriteLine($"  - '{content}' (fill_type: {fillType})");
                    }
                }
                
                if (gradientShapeFills.Count > 0)
                {
                    _output.WriteLine("\nAll gradient shape fills:");
                    foreach (var (shapeName, fillType) in gradientShapeFills)
                    {
                        _output.WriteLine($"  - '{shapeName}' (fill_type: {fillType})");
                    }
                }

                if (gradientTextRuns.Count == 0 && gradientShapeFills.Count == 0)
                {
                    _output.WriteLine("\n⚠ No gradient text runs or gradient shape fills found in the JSON output.");
                    _output.WriteLine("This may indicate:");
                    _output.WriteLine("  1. The PPTX file doesn't contain gradient fill text or gradient-filled shapes");
                    _output.WriteLine("  2. Gradient fill extraction is not working correctly");
                    _output.WriteLine("  3. The text content may have changed");
                }
                
                // Assert that we found some gradient information (either text or shape fill)
                Assert.True(gradientTextCount > 0 || gradientShapeFillCount > 0, 
                    $"Should find at least one gradient text run or gradient-filled shape. Total text runs checked: {totalTextRuns}");
            }
        }
    }
}
