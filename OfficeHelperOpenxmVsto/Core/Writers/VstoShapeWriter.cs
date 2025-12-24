using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using OfficeHelperOpenXml.Models.Json;
using OfficeHelperOpenXml.Utils;
// 使用别名避免 Shape 类型冲突（Microsoft.Office.Core 和 Microsoft.Office.Interop.PowerPoint 都有 Shape 类型）
using PptShape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace OfficeHelperOpenXml.Core.Writers
{
    /// <summary>
    /// VSTO 形状写入器
    /// </summary>
    public class VstoShapeWriter
    {
        private readonly VstoStyleWriter _styleWriter;

        public VstoShapeWriter()
        {
            _styleWriter = new VstoStyleWriter();
        }

        /// <summary>
        /// 创建形状
        /// </summary>
        public PptShape CreateShape(Slide slide, ShapeJsonData shapeData)
        {
            if (slide == null || shapeData == null) return null;

            try
            {
                PptShape shape = null;

                // 根据类型创建不同的形状
                switch (shapeData.Type?.ToLower())
                {
                    case "textbox":
                        shape = CreateTextBox(slide, shapeData);
                        break;
                    case "autoshape":
                        shape = CreateAutoShape(slide, shapeData);
                        break;
                    case "picture":
                        shape = CreatePicture(slide, shapeData);
                        break;
                    case "table":
                        shape = CreateTable(slide, shapeData);
                        break;
                    case "group":
                        shape = CreateGroup(slide, shapeData);
                        break;
                    case "connection":
                        shape = CreateConnection(slide, shapeData);
                        break;
                    default:
                        // 默认创建文本框
                        shape = CreateTextBox(slide, shapeData);
                        break;
                }

                if (shape != null)
                {
                    // 应用位置和大小
                    ApplyPositionAndSize(shape, shapeData);
                    
                    // 应用旋转
                    if (shapeData.Rotation != 0)
                    {
                        shape.Rotation = shapeData.Rotation;
                    }
                    
                    // 应用样式
                    _styleWriter.ApplyFill(shape, shapeData.Fill);
                    _styleWriter.ApplyLine(shape, shapeData.Line);
                    _styleWriter.ApplyShadow(shape, shapeData.Shadow);
                }

                return shape;
            }
            catch (Exception ex)
            {
                var logger = new Logger();
                logger.LogError($"创建形状失败: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 创建文本框
        /// </summary>
        private PptShape CreateTextBox(Slide slide, ShapeJsonData shapeData)
        {
            if (!shapeData.TryParseBox(out float left, out float top, out float width, out float height))
                return null;

            float leftPt = VstoHelper.CmToPoints(left);
            float topPt = VstoHelper.CmToPoints(top);
            float widthPt = VstoHelper.CmToPoints(width);
            float heightPt = VstoHelper.CmToPoints(height);

            var shape = slide.Shapes.AddTextbox(
                MsoTextOrientation.msoTextOrientationHorizontal,
                leftPt, topPt, widthPt, heightPt);

            shape.Name = string.IsNullOrEmpty(shapeData.Name) ? "TextBox" : shapeData.Name;

            // 应用文本
            if (shapeData.HasText == 1 && shapeData.Text != null && shapeData.Text.Count > 0)
            {
                ApplyText(shape, shapeData.Text);
            }

            return shape;
        }

        /// <summary>
        /// 创建自动形状
        /// </summary>
        private PptShape CreateAutoShape(Slide slide, ShapeJsonData shapeData)
        {
            if (!shapeData.TryParseBox(out float left, out float top, out float width, out float height))
                return null;

            float leftPt = VstoHelper.CmToPoints(left);
            float topPt = VstoHelper.CmToPoints(top);
            float widthPt = VstoHelper.CmToPoints(width);
            float heightPt = VstoHelper.CmToPoints(height);

            // 根据 special_type 确定形状类型，默认为矩形
            MsoAutoShapeType shapeType = MsoAutoShapeType.msoShapeRectangle;
            
            if (!string.IsNullOrEmpty(shapeData.SpecialType))
            {
                shapeType = GetAutoShapeType(shapeData.SpecialType);
            }

            var shape = slide.Shapes.AddShape(
                shapeType,
                leftPt, topPt, widthPt, heightPt);

            shape.Name = string.IsNullOrEmpty(shapeData.Name) ? "AutoShape" : shapeData.Name;

            // 应用文本
            if (shapeData.HasText == 1 && shapeData.Text != null && shapeData.Text.Count > 0)
            {
                ApplyText(shape, shapeData.Text);
            }

            return shape;
        }

        /// <summary>
        /// 创建图片（占位符，实际需要图片路径）
        /// </summary>
        private PptShape CreatePicture(Slide slide, ShapeJsonData shapeData)
        {
            if (!shapeData.TryParseBox(out float left, out float top, out float width, out float height))
                return null;

            float leftPt = VstoHelper.CmToPoints(left);
            float topPt = VstoHelper.CmToPoints(top);
            float widthPt = VstoHelper.CmToPoints(width);
            float heightPt = VstoHelper.CmToPoints(height);

            // 注意：实际图片插入需要图片文件路径
            // 这里创建一个占位符矩形
            var shape = slide.Shapes.AddShape(
                MsoAutoShapeType.msoShapeRectangle,
                leftPt, topPt, widthPt, heightPt);

            shape.Name = string.IsNullOrEmpty(shapeData.Name) ? "Picture" : shapeData.Name;
            shape.Fill.Visible = MsoTriState.msoFalse;
            shape.Line.DashStyle = MsoLineDashStyle.msoLineDash;

            return shape;
        }

        /// <summary>
        /// 创建表格（简化实现）
        /// </summary>
        private PptShape CreateTable(Slide slide, ShapeJsonData shapeData)
        {
            if (!shapeData.TryParseBox(out float left, out float top, out float width, out float height))
                return null;

            float leftPt = VstoHelper.CmToPoints(left);
            float topPt = VstoHelper.CmToPoints(top);
            float widthPt = VstoHelper.CmToPoints(width);
            float heightPt = VstoHelper.CmToPoints(height);

            // 创建简单的表格（2x2 作为默认）
            var shape = slide.Shapes.AddTable(2, 2, leftPt, topPt, widthPt, heightPt);
            shape.Name = string.IsNullOrEmpty(shapeData.Name) ? "Table" : shapeData.Name;

            return shape;
        }

        /// <summary>
        /// 创建组合形状
        /// </summary>
        private PptShape CreateGroup(Slide slide, ShapeJsonData shapeData)
        {
            // 组合形状需要先创建子形状，然后组合
            // 这里简化处理，创建一个矩形作为占位符
            if (!shapeData.TryParseBox(out float left, out float top, out float width, out float height))
                return null;

            float leftPt = VstoHelper.CmToPoints(left);
            float topPt = VstoHelper.CmToPoints(top);
            float widthPt = VstoHelper.CmToPoints(width);
            float heightPt = VstoHelper.CmToPoints(height);

            var shape = slide.Shapes.AddShape(
                MsoAutoShapeType.msoShapeRectangle,
                leftPt, topPt, widthPt, heightPt);

            shape.Name = string.IsNullOrEmpty(shapeData.Name) ? "Group" : shapeData.Name;

            // TODO: 如果有 children 数据，创建子形状并组合

            return shape;
        }

        /// <summary>
        /// 创建连接线
        /// </summary>
        private PptShape CreateConnection(Slide slide, ShapeJsonData shapeData)
        {
            if (!shapeData.TryParseBox(out float left, out float top, out float width, out float height))
                return null;

            float leftPt = VstoHelper.CmToPoints(left);
            float topPt = VstoHelper.CmToPoints(top);
            float widthPt = VstoHelper.CmToPoints(width);
            float heightPt = VstoHelper.CmToPoints(height);

            // 创建连接线
            var shape = slide.Shapes.AddConnector(
                MsoConnectorType.msoConnectorStraight,
                leftPt, topPt, leftPt + widthPt, topPt + heightPt);

            shape.Name = string.IsNullOrEmpty(shapeData.Name) ? "Connection" : shapeData.Name;

            return shape;
        }

        /// <summary>
        /// 应用位置和大小
        /// </summary>
        private void ApplyPositionAndSize(PptShape shape, ShapeJsonData shapeData)
        {
            if (shape == null || shapeData == null) return;

            if (shapeData.TryParseBox(out float left, out float top, out float width, out float height))
            {
                shape.Left = VstoHelper.CmToPoints(left);
                shape.Top = VstoHelper.CmToPoints(top);
                shape.Width = VstoHelper.CmToPoints(width);
                shape.Height = VstoHelper.CmToPoints(height);
            }
        }

        /// <summary>
        /// 应用文本
        /// </summary>
        private void ApplyText(PptShape shape, List<TextRunJsonData> textRuns)
        {
            if (shape == null || textRuns == null || textRuns.Count == 0) return;

            try
            {
                var textFrame = shape.TextFrame;
                if (textFrame == null) return;

                var textRange = textFrame.TextRange;
                if (textRange == null) return;

                // 清除现有文本
                textRange.Text = "";

                // 应用每个文本运行
                foreach (var textRun in textRuns)
                {
                    if (textRun == null) continue;

                    // 添加文本内容
                    var newRange = textRange.InsertAfter(textRun.Content ?? "");
                    
                    // 应用文本样式
                    if (!string.IsNullOrEmpty(textRun.Content))
                    {
                        _styleWriter.ApplyTextFormat(newRange, textRun);
                    }
                }
            }
            catch (Exception ex)
            {
                var logger = new Logger();
                logger.LogWarning($"应用文本失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 根据特殊类型字符串获取自动形状类型
        /// </summary>
        private MsoAutoShapeType GetAutoShapeType(string specialType)
        {
            // 常见形状类型映射
            switch (specialType.ToLower())
            {
                case "矩形":
                case "rectangle":
                    return MsoAutoShapeType.msoShapeRectangle;
                case "椭圆":
                case "ellipse":
                case "圆形":
                case "circle":
                    return MsoAutoShapeType.msoShapeOval;
                case "圆角矩形":
                case "rounded rectangle":
                    return MsoAutoShapeType.msoShapeRoundedRectangle;
                case "三角形":
                case "triangle":
                    return MsoAutoShapeType.msoShapeIsoscelesTriangle;
                case "菱形":
                case "diamond":
                    return MsoAutoShapeType.msoShapeDiamond;
                case "箭头":
                case "arrow":
                    return MsoAutoShapeType.msoShapeRightArrow;
                default:
                    return MsoAutoShapeType.msoShapeRectangle;
            }
        }
    }
}

