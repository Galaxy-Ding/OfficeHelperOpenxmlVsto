using OfficeHelperOpenXml.Api.PowerPoint;

namespace OfficeHelperOpenXml.Api.PowerPoint
{
    /// <summary>
    /// PowerPoint 写入器工厂类
    /// </summary>
    public static class PowerPointWriterFactory
    {
        /// <summary>
        /// 创建 PowerPoint 写入器实例
        /// </summary>
        public static IPowerPointWriter CreateWriter()
        {
            return new PowerPointWriter();
        }
    }
}

