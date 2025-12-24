using System;
using OfficeHelperOpenXml.Models.Json;

namespace OfficeHelperOpenXml.Api.PowerPoint
{
    /// <summary>
    /// PowerPoint 写入器接口
    /// </summary>
    public interface IPowerPointWriter : IDisposable
    {
        /// <summary>
        /// 从模板文件打开或创建演示文稿
        /// </summary>
        /// <param name="templatePath">模板文件路径</param>
        /// <returns>是否成功</returns>
        bool OpenFromTemplate(string templatePath);

        /// <summary>
        /// 清除所有内容幻灯片的内容（保留母版）
        /// </summary>
        /// <returns>是否成功</returns>
        bool ClearAllContentSlides();

        /// <summary>
        /// 从 JSON 字符串写入内容
        /// </summary>
        /// <param name="jsonData">JSON 数据字符串</param>
        /// <returns>是否成功</returns>
        bool WriteFromJson(string jsonData);

        /// <summary>
        /// 从 PresentationJsonData 对象写入内容
        /// </summary>
        /// <param name="jsonData">JSON 数据对象</param>
        /// <returns>是否成功</returns>
        bool WriteFromJsonData(PresentationJsonData jsonData);

        /// <summary>
        /// 保存到文件
        /// </summary>
        /// <param name="outputPath">输出文件路径</param>
        /// <returns>是否成功</returns>
        bool SaveAs(string outputPath);

        /// <summary>
        /// 关闭文档
        /// </summary>
        void Close();
    }
}

