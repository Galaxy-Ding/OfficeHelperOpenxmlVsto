using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using OfficeHelperOpenXml.Models.Json;
using OfficeHelperOpenXml.Utils;

namespace OfficeHelperOpenXml.Core.Writers
{
    /// <summary>
    /// VSTO 幻灯片写入器
    /// </summary>
    public class VstoSlideWriter
    {
        private readonly VstoShapeWriter _shapeWriter;
        private readonly Presentation _presentation;

        public VstoSlideWriter(Presentation presentation)
        {
            _presentation = presentation ?? throw new ArgumentNullException(nameof(presentation));
            _shapeWriter = new VstoShapeWriter();
        }

        /// <summary>
        /// 获取或创建幻灯片
        /// </summary>
        public Slide GetOrCreateSlide(int slideIndex)
        {
            if (_presentation == null) return null;

            try
            {
                // slideIndex 是 1-based（VSTO 使用 1-based 索引）
                int vstoIndex = slideIndex + 1;

                // 如果幻灯片已存在，返回它
                if (vstoIndex <= _presentation.Slides.Count)
                {
                    return _presentation.Slides[vstoIndex];
                }

                // 否则创建新幻灯片
                // 使用第一个母版布局（通常是标题和内容布局）
                var layout = _presentation.SlideMaster.CustomLayouts[1];
                var slide = _presentation.Slides.AddSlide(vstoIndex, layout);

                return slide;
            }
            catch (Exception ex)
            {
                var logger = new Logger();
                logger.LogError($"获取或创建幻灯片失败: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 清除幻灯片内容
        /// </summary>
        public void ClearSlideContent(Slide slide)
        {
            if (slide == null) return;

            try
            {
                int shapeCount = slide.Shapes.Count;
                int deletedCount = 0;
                int skippedCount = 0;

                // 从后往前删除，避免索引问题
                for (int i = shapeCount; i >= 1; i--)
                {
                    try
                    {
                        var shape = slide.Shapes[i];
                        
                        try
                        {
                            // 尝试删除形状
                            // 如果形状在母版上，Delete() 会失败或抛出异常
                            // 我们通过捕获异常来处理这种情况
                            shape.Delete();
                            deletedCount++;
                        }
                        catch (System.Runtime.InteropServices.COMException)
                        {
                            // 如果删除失败（可能是母版形状或锁定形状），跳过
                            skippedCount++;
                        }
                        catch
                        {
                            // 其他异常也跳过
                            skippedCount++;
                        }
                    }
                    catch (Exception ex)
                    {
                        var logger = new Logger();
                        logger.LogWarning($"删除形状时出错: {ex.Message}");
                        // 继续处理下一个形状
                    }
                }

                var logger2 = new Logger();
                logger2.LogSuccess($"清除幻灯片内容完成: 删除 {deletedCount} 个形状，跳过 {skippedCount} 个形状");
            }
            catch (Exception ex)
            {
                var logger = new Logger();
                logger.LogWarning($"清除幻灯片内容失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 写入幻灯片数据
        /// </summary>
        public void WriteSlideData(Slide slide, SlideJsonData slideData)
        {
            if (slide == null || slideData == null) return;

            try
            {
                // 清除现有内容
                ClearSlideContent(slide);

                // 写入形状
                if (slideData.Shapes != null && slideData.Shapes.Count > 0)
                {
                    foreach (var shapeData in slideData.Shapes)
                    {
                        if (shapeData != null)
                        {
                            _shapeWriter.CreateShape(slide, shapeData);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                var logger = new Logger();
                logger.LogError($"写入幻灯片数据失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 清除所有内容幻灯片的内容
        /// </summary>
        public void ClearAllContentSlides()
        {
            if (_presentation == null) return;

            try
            {
                int slideCount = _presentation.Slides.Count;
                var logger = new Logger();
                logger.LogSuccess($"开始清除 {slideCount} 张幻灯片的内容");

                // 从后往前遍历，避免索引问题
                for (int i = slideCount; i >= 1; i--)
                {
                    try
                    {
                        var slide = _presentation.Slides[i];
                        ClearSlideContent(slide);
                    }
                    catch (Exception ex)
                    {
                        var logger2 = new Logger();
                        logger2.LogWarning($"清除第 {i} 张幻灯片失败: {ex.Message}");
                        // 继续处理下一张幻灯片
                    }
                }

                logger.LogSuccess($"已清除所有 {slideCount} 张幻灯片的内容");
            }
            catch (Exception ex)
            {
                var logger = new Logger();
                logger.LogError($"清除所有内容幻灯片失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 写入多个幻灯片
        /// </summary>
        public void WriteSlides(List<SlideJsonData> slidesData)
        {
            if (slidesData == null || slidesData.Count == 0) return;

            if (_presentation == null) return;

            try
            {
                // 注意：PowerPoint 的 Application 对象可能不支持 ScreenUpdating 属性
                // 这里我们直接进行批量写入，通过减少日志输出等方式优化性能

                var logger = new Logger();
                logger.LogSuccess($"开始写入 {slidesData.Count} 张幻灯片");

                for (int i = 0; i < slidesData.Count; i++)
                {
                    try
                    {
                        var slideData = slidesData[i];
                        if (slideData == null) continue;

                        // page_number 是 1-based，转换为 0-based 索引
                        int slideIndex = slideData.PageNumber > 0 ? slideData.PageNumber - 1 : i;
                        
                        var slide = GetOrCreateSlide(slideIndex);
                        if (slide != null)
                        {
                            WriteSlideData(slide, slideData);
                        }
                    }
                    catch (Exception ex)
                    {
                        var logger2 = new Logger();
                        logger2.LogWarning($"写入第 {i + 1} 张幻灯片失败: {ex.Message}");
                        // 继续处理下一张幻灯片
                    }
                }

                logger.LogSuccess($"成功写入 {slidesData.Count} 张幻灯片");
            }
            catch (Exception ex)
            {
                var logger = new Logger();
                logger.LogError($"写入多个幻灯片失败: {ex.Message}");
            }
        }
    }
}

