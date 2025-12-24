using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeHelperOpenXml.Elements;
using OfficeHelperOpenXml.Models;
using OfficeHelperOpenXml.Utils;

namespace OfficeHelperOpenXml.Core.Readers
{
    public class PresentationReader
    {
        private readonly SlideReader _slideReader;
        private readonly StyleReader _styleReader;
        
        public PresentationReader()
        {
            _slideReader = new SlideReader();
            _styleReader = new StyleReader();
        }
        
        public PresentationInfo ReadPresentation(string filePath)
        {
            var info = new PresentationInfo();
            
            try
            {
                using (var doc = PresentationDocument.Open(filePath, false))
                {
                    var presentationPart = doc.PresentationPart;
                    if (presentationPart == null)
                    {
                        info.Error = "无法读取演示文稿";
                        return info;
                    }
                    
                    // 读取演示文稿尺寸
                    var slideSize = presentationPart.Presentation?.SlideSize;
                    if (slideSize != null)
                    {
                        info.SlideWidth = (float)UnitConverter.EmuToCm(slideSize.Cx?.Value ?? 0);
                        info.SlideHeight = (float)UnitConverter.EmuToCm(slideSize.Cy?.Value ?? 0);
                    }
                    
                    // 提取默认文本样式
                    Console.WriteLine("[DEBUG] ========== Starting Style Extraction ==========");
                    info.DefaultTextStyle = _styleReader.ExtractDefaultTextStyle(presentationPart);
                    
                    // 提取母版样式
                    info.SlideMasterStyles = ReadSlideMasterStyles(presentationPart);
                    Console.WriteLine($"[DEBUG] Total masters extracted: {info.SlideMasterStyles?.Count ?? 0}");
                    Console.WriteLine("[DEBUG] ========== Style Extraction Complete ==========");
                    
                    // 读取母版和布局（保留用于形状提取）
                    info.MasterSlides = ReadMasterSlides(presentationPart);
                    
                    // 读取所有幻灯片
                    var slideIdList = presentationPart.Presentation?.SlideIdList;
                    if (slideIdList != null)
                    {
                        int slideIndex = 0;
                        foreach (var slideId in slideIdList.Elements<SlideId>())
                        {
                            try
                            {
                                var slidePart = presentationPart.GetPartById(slideId.RelationshipId) as SlidePart;
                                if (slidePart != null)
                                {
                                    var slideInfo = new SlideInfo
                                    {
                                        SlideIndex = slideIndex,
                                        SlideNumber = slideIndex + 1,
                                        Elements = _slideReader.ReadSlide(slidePart)
                                    };
                                    info.Slides.Add(slideInfo);
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"读取幻灯片 {slideIndex + 1} 时出错: {ex.Message}");
                            }
                            slideIndex++;
                        }
                    }
                    
                    info.SlideCount = info.Slides.Count;
                }
            }
            catch (Exception ex)
            {
                info.Error = $"读取演示文稿时出错: {ex.Message}";
            }
            
            return info;
        }
        
        private List<MasterSlideInfo> ReadMasterSlides(PresentationPart presentationPart)
        {
            var masterSlides = new List<MasterSlideInfo>();
            int pageNumber = 0;
            
            try
            {
                foreach (var masterPart in presentationPart.SlideMasterParts)
                {
                    // 1. 添加母版本身 (page_number = 0)
                    var masterName = masterPart.SlideMaster?.CommonSlideData?.Name?.Value ?? "幻灯片母版";
                    var masterInfo = new MasterSlideInfo
                    {
                        PageNumber = pageNumber++,
                        Title = masterName,
                        SubTitle = "",
                        Shapes = _slideReader.ReadSlide(masterPart),
                        IsMaster = true
                    };
                    masterSlides.Add(masterInfo);
                    
                    // 2. 添加所有布局 (page_number = 1, 2, 3, ...)
                    foreach (var layoutPart in masterPart.SlideLayoutParts)
                    {
                        var layoutName = layoutPart.SlideLayout?.CommonSlideData?.Name?.Value ?? "布局";
                        var layoutInfo = new MasterSlideInfo
                        {
                            PageNumber = pageNumber++,
                            Title = $"{layoutName}({masterName})",
                            SubTitle = "",
                            Shapes = _slideReader.ReadSlide(layoutPart),
                            IsMaster = false
                        };
                        masterSlides.Add(layoutInfo);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"读取母版时出错: {ex.Message}");
            }
            
            return masterSlides;
        }
        
        private List<Models.Json.SlideMasterStyleJsonData> ReadSlideMasterStyles(PresentationPart presentationPart)
        {
            var masterStyles = new List<Models.Json.SlideMasterStyleJsonData>();
            
            try
            {
                uint masterIdCounter = 2147483648; // OpenXML 默认起始 ID
                
                foreach (var masterPart in presentationPart.SlideMasterParts)
                {
                    var masterStyle = _styleReader.ExtractSlideMasterStyle(masterPart, masterIdCounter);
                    masterStyles.Add(masterStyle);
                    masterIdCounter++;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"读取母版样式时出错: {ex.Message}");
            }
            
            return masterStyles;
        }
    }
    
    public class PresentationInfo
    {
        public float SlideWidth { get; set; }
        public float SlideHeight { get; set; }
        public int SlideCount { get; set; }
        public List<SlideInfo> Slides { get; set; } = new List<SlideInfo>();
        public List<MasterSlideInfo> MasterSlides { get; set; } = new List<MasterSlideInfo>();
        
        // 新增：样式信息
        public Models.Json.DefaultTextStyleJsonData DefaultTextStyle { get; set; } = new Models.Json.DefaultTextStyleJsonData();
        public List<Models.Json.SlideMasterStyleJsonData> SlideMasterStyles { get; set; } = new List<Models.Json.SlideMasterStyleJsonData>();
        
        public string Error { get; set; }
    }
    
    public class SlideInfo
    {
        public int SlideIndex { get; set; }
        public int SlideNumber { get; set; }
        public List<BaseElement> Elements { get; set; } = new List<BaseElement>();
    }
    
    public class MasterSlideInfo
    {
        public int PageNumber { get; set; }
        public string Title { get; set; }
        public string SubTitle { get; set; }
        public List<BaseElement> Shapes { get; set; } = new List<BaseElement>();
        public bool IsMaster { get; set; } // true=母版本身, false=布局
    }
}
