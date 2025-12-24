using System;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using OfficeHelperOpenXml.Core.Converters;
using OfficeHelperOpenXml.Core.Writers;
using OfficeHelperOpenXml.Models.Json;
using OfficeHelperOpenXml.Utils;

namespace OfficeHelperOpenXml.Api.PowerPoint
{
    /// <summary>
    /// PowerPoint 写入器实现
    /// </summary>
    public class PowerPointWriter : IPowerPointWriter
    {
        private Application _app;
        private Presentation _presentation;
        private VstoSlideWriter _slideWriter;
        private JsonToVstoConverter _converter;
        private bool _disposed = false;
        private bool _appCreatedByUs = false;  // 标记是否是我们创建的 PowerPoint 实例

        /// <summary>
        /// 从模板文件打开或创建演示文稿
        /// </summary>
        public bool OpenFromTemplate(string templatePath)
        {
            var logger = new Logger();
            
            if (string.IsNullOrEmpty(templatePath))
            {
                logger.LogError("模板文件路径不能为空");
                return false;
            }

            // 检查文件是否存在
            if (!File.Exists(templatePath))
            {
                logger.LogError($"模板文件不存在: {templatePath}");
                return false;
            }

            // 检查文件是否可读
            try
            {
                using (var stream = File.OpenRead(templatePath))
                {
                    // 文件可读
                }
            }
            catch (UnauthorizedAccessException)
            {
                logger.LogError($"模板文件无访问权限: {templatePath}");
                return false;
            }
            catch (IOException ex)
            {
                logger.LogError($"无法读取模板文件: {ex.Message}");
                return false;
            }

            try
            {
                // 检查 PowerPoint 是否可用
                if (!VstoHelper.IsPowerPointAvailable())
                {
                    logger.LogError("PowerPoint 不可用，请确保已安装 Microsoft PowerPoint");
                    return false;
                }

                // ⭐ 策略1：智能实例管理 - 尝试获取现有的 PowerPoint 实例
                try
                {
                    _app = (Application)Marshal.GetActiveObject("PowerPoint.Application");
                    _appCreatedByUs = false;  // 连接到现有实例
                    logger.LogInfo("已连接到现有的 PowerPoint 实例");
                }
                catch (COMException)
                {
                    // 没有现有实例，创建新实例
                    _app = new Application();
                    _appCreatedByUs = true;  // 标记为我们创建的实例
                    logger.LogInfo("创建了新的 PowerPoint 实例");
                    
                    // 尝试隐藏窗口（某些版本的 PowerPoint 可能不支持，如果失败就继续执行）
                    try
                    {
                        _app.Visible = MsoTriState.msoFalse; // 后台运行
                    }
                    catch (COMException)
                    {
                        // 某些版本的 PowerPoint 不允许隐藏窗口，忽略此错误继续执行
                        // 窗口将保持可见，但不影响功能
                    }
                }
                
                _app.DisplayAlerts = PpAlertLevel.ppAlertsNone; // 不显示警告

                // 打开模板文件（使用绝对路径）
                // PowerPoint.Presentations.Open(string FileName,
                //                               MsoTriState ReadOnly,
                //                               MsoTriState Untitled,
                //                               MsoTriState WithWindow)
                string absolutePath = Path.GetFullPath(templatePath);
                _presentation = _app.Presentations.Open(
                    absolutePath,
                    ReadOnly: MsoTriState.msoTrue,
                    Untitled: MsoTriState.msoFalse,
                    WithWindow: MsoTriState.msoFalse);

                if (_presentation == null)
                {
                    logger.LogError("打开模板文件失败：返回 null");
                    Cleanup();
                    return false;
                }

                // 初始化写入器
                _slideWriter = new VstoSlideWriter(_presentation);
                _converter = new JsonToVstoConverter();

                logger.LogSuccess($"成功打开模板文件: {templatePath}");
                return true;
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                logger.LogError($"COM 错误：打开模板文件失败: {ex.Message} (HRESULT: 0x{ex.ErrorCode:X})");
                Cleanup();
                return false;
            }
            catch (FileNotFoundException ex)
            {
                logger.LogError($"文件未找到: {ex.Message}");
                Cleanup();
                return false;
            }
            catch (UnauthorizedAccessException ex)
            {
                logger.LogError($"访问被拒绝: {ex.Message}");
                Cleanup();
                return false;
            }
            catch (Exception ex)
            {
                logger.LogError($"打开模板文件时出错: {ex.Message}\n堆栈跟踪: {ex.StackTrace}");
                Cleanup();
                return false;
            }
        }

        /// <summary>
        /// 清除所有内容幻灯片的内容（保留母版）
        /// </summary>
        public bool ClearAllContentSlides()
        {
            var logger = new Logger();
            
            if (_presentation == null)
            {
                logger.LogError("演示文稿未打开");
                return false;
            }

            try
            {
                _slideWriter?.ClearAllContentSlides();
                logger.LogSuccess("已清除所有内容幻灯片");
                return true;
            }
            catch (Exception ex)
            {
                logger.LogError($"清除内容幻灯片失败: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 从 JSON 字符串写入内容
        /// </summary>
        public bool WriteFromJson(string jsonData)
        {
            var logger = new Logger();
            
            if (string.IsNullOrEmpty(jsonData))
            {
                logger.LogError("JSON 数据不能为空");
                return false;
            }

            try
            {
                var presentationData = _converter?.ParseJson(jsonData);
                if (presentationData == null)
                {
                    logger.LogError("JSON 解析失败");
                    return false;
                }

                return WriteFromJsonData(presentationData);
            }
            catch (Exception ex)
            {
                logger.LogError($"从 JSON 写入内容失败: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 从 PresentationJsonData 对象写入内容
        /// </summary>
        public bool WriteFromJsonData(PresentationJsonData jsonData)
        {
            var logger = new Logger();
            
            if (jsonData == null)
            {
                logger.LogError("JSON 数据对象不能为空");
                return false;
            }

            if (_presentation == null || _slideWriter == null)
            {
                logger.LogError("演示文稿未打开");
                return false;
            }

            try
            {
                // 写入内容幻灯片
                if (jsonData.ContentSlides != null && jsonData.ContentSlides.Count > 0)
                {
                    _slideWriter.WriteSlides(jsonData.ContentSlides);
                    logger.LogSuccess($"成功写入 {jsonData.ContentSlides.Count} 张内容幻灯片");
                }
                else
                {
                    logger.LogWarning("没有内容幻灯片需要写入");
                }

                return true;
            }
            catch (Exception ex)
            {
                logger.LogError($"写入内容失败: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 保存到文件
        /// </summary>
        public bool SaveAs(string outputPath)
        {
            var logger = new Logger();
            
            if (string.IsNullOrEmpty(outputPath))
            {
                logger.LogError("输出文件路径不能为空");
                return false;
            }

            if (_presentation == null)
            {
                logger.LogError("演示文稿未打开");
                return false;
            }

            try
            {
                // 确保输出目录存在
                var directory = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    try
                    {
                        Directory.CreateDirectory(directory);
                    }
                    catch (Exception ex)
                    {
                        logger.LogError($"创建输出目录失败: {ex.Message}");
                        return false;
                    }
                }

                // 使用绝对路径
                string absolutePath = Path.GetFullPath(outputPath);
                
                // 如果文件已存在，使用临时文件策略避免文件占用问题
                string tempPath = null;
                bool useTempFile = File.Exists(absolutePath);
                
                if (useTempFile)
                {
                    // 生成临时文件路径
                    tempPath = Path.Combine(
                        directory ?? Path.GetTempPath(),
                        Path.GetFileNameWithoutExtension(absolutePath) + "_temp_" + Guid.NewGuid().ToString("N").Substring(0, 8) + Path.GetExtension(absolutePath)
                    );
                    logger.LogInfo($"[SaveAs] 文件已存在，将使用临时文件保存: {tempPath}");
                }

                // ⭐ 关键修复：使用 SaveCopyAs 而不是 SaveAs
                // SaveCopyAs 保存副本而不影响当前打开的文件，可以避免文件句柄占用问题
                string savePath = useTempFile ? tempPath : absolutePath;
                logger.LogInfo($"[SaveAs] 准备调用 _presentation.SaveCopyAs，路径: {savePath}");
                
                try
                {
                    // 使用 SaveCopyAs 保存副本，这样不会影响当前打开的文件
                    _presentation.SaveCopyAs(savePath, PpSaveAsFileType.ppSaveAsDefault);
                    logger.LogInfo("[SaveAs] _presentation.SaveCopyAs 调用返回，未抛出异常");
                }
                catch (System.Runtime.InteropServices.COMException ex)
                {
                    // 如果 SaveCopyAs 失败，回退到 SaveAs
                    logger.LogWarning($"SaveCopyAs 失败，尝试使用 SaveAs: {ex.Message}");
                    try
                    {
                        _presentation.SaveAs(savePath, PpSaveAsFileType.ppSaveAsDefault, MsoTriState.msoTrue);
                        logger.LogInfo("[SaveAs] _presentation.SaveAs 调用返回，未抛出异常");
                    }
                    catch (System.Runtime.InteropServices.COMException ex2)
                    {
                        // 如果保存失败，尝试等待并重试（最多3次）
                        if (ex2.ErrorCode == unchecked((int)0x80004005)) // E_FAIL - 文件可能被占用
                        {
                            logger.LogWarning($"首次保存失败（文件可能被占用），将等待后重试...");
                            for (int retry = 1; retry <= 3; retry++)
                            {
                                System.Threading.Thread.Sleep(500 * retry); // 递增等待时间：500ms, 1000ms, 1500ms
                                try
                                {
                                    logger.LogInfo($"[SaveAs] 第 {retry} 次重试保存...");
                                    _presentation.SaveAs(savePath, PpSaveAsFileType.ppSaveAsDefault, MsoTriState.msoTrue);
                                    logger.LogInfo("[SaveAs] 重试保存成功");
                                    break;
                                }
                                catch
                                {
                                    if (retry == 3)
                                    {
                                        throw; // 最后一次重试失败，抛出异常
                                    }
                                }
                            }
                        }
                        else
                        {
                            throw; // 其他 COM 错误直接抛出
                        }
                    }
                }
                
                // ⭐ 关键修复：等待 PowerPoint 完全释放文件句柄
                // PowerPoint 在保存后可能仍持有文件句柄，需要等待一段时间
                if (useTempFile)
                {
                    logger.LogInfo("[SaveAs] 等待 PowerPoint 释放文件句柄...");
                    System.Threading.Thread.Sleep(1500); // 等待 1.5 秒
                    
                    // 尝试强制刷新 PowerPoint 的文件操作
                    try
                    {
                        // 通过访问演示文稿属性来确保保存操作完成
                        var _ = _presentation.FullName;
                        System.Threading.Thread.Sleep(500); // 再等待 0.5 秒
                    }
                    catch
                    {
                        // 忽略错误，继续执行
                    }
                }
                
                // 如果使用临时文件，现在替换原文件
                if (useTempFile && !string.IsNullOrEmpty(tempPath) && File.Exists(tempPath))
                {
                    logger.LogInfo($"[SaveAs] 准备将临时文件替换原文件...");
                    
                    // ⭐ 关键修复：在删除原文件前，检查文件锁定情况
                    if (File.Exists(absolutePath))
                    {
                        try
                        {
                            // 尝试以独占方式打开文件，检测是否被锁定
                            using (var fs = File.Open(absolutePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                            {
                                // 文件未被锁定，可以删除
                            }
                        }
                        catch (IOException)
                        {
                            // 文件被锁定，记录详细信息
                            logger.LogWarning($"原文件被锁定，正在检查占用进程...");
                            string lockInfo = FileLockChecker.CheckFileLock(absolutePath);
                            logger.LogInfo($"文件锁定检查结果:\n{lockInfo}");
                            
                            // 等待更长时间，让 PowerPoint 释放文件
                            logger.LogInfo("[SaveAs] 等待 PowerPoint 释放原文件句柄（额外等待 2 秒）...");
                            System.Threading.Thread.Sleep(2000);
                            
                            // 强制垃圾回收
                            System.GC.Collect();
                            System.GC.WaitForPendingFinalizers();
                            System.GC.Collect();
                        }
                    }
                    
                    // 尝试删除原文件（最多重试5次，增加重试次数）
                    int maxRetries = 5;
                    for (int retry = 0; retry < maxRetries; retry++)
                    {
                        try
                        {
                            if (File.Exists(absolutePath))
                            {
                                File.Delete(absolutePath);
                                logger.LogInfo($"[SaveAs] 原文件删除成功（第 {retry + 1} 次尝试）");
                            }
                            break; // 删除成功，退出循环
                        }
                        catch (Exception ex)
                        {
                            if (retry < maxRetries - 1)
                            {
                                int waitTime = 1000 * (retry + 1); // 递增等待时间：1s, 2s, 3s, 4s, 5s
                                logger.LogWarning($"删除原文件失败（第 {retry + 1} 次尝试）: {ex.Message}，等待 {waitTime}ms 后重试...");
                                
                                // 每次重试前再次检查文件锁定
                                if (File.Exists(absolutePath))
                                {
                                    string lockInfo = FileLockChecker.CheckFileLock(absolutePath);
                                    logger.LogInfo($"重试前文件锁定检查:\n{lockInfo}");
                                }
                                
                                System.Threading.Thread.Sleep(waitTime);
                                
                                // 强制垃圾回收
                                System.GC.Collect();
                                System.GC.WaitForPendingFinalizers();
                                System.GC.Collect();
                            }
                            else
                            {
                                logger.LogError($"无法删除原文件: {ex.Message}");
                                // 最后一次尝试：输出详细的文件锁定信息
                                if (File.Exists(absolutePath))
                                {
                                    string lockInfo = FileLockChecker.CheckFileLock(absolutePath);
                                    logger.LogError($"最终文件锁定检查结果:\n{lockInfo}");
                                }
                                // 删除临时文件
                                try { File.Delete(tempPath); } catch { }
                                return false;
                            }
                        }
                    }
                    
                    // ⭐ 关键修复：移动临时文件到目标位置（带重试机制）
                    int moveRetries = 5; // 增加重试次数
                    bool moveSuccess = false;
                    for (int retry = 0; retry < moveRetries; retry++)
                    {
                        try
                        {
                            // 检查临时文件是否可访问
                            using (var fs = File.Open(tempPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                            {
                                // 文件可访问，继续
                            }
                            
                            File.Move(tempPath, absolutePath);
                            logger.LogInfo($"[SaveAs] 临时文件已成功替换原文件");
                            moveSuccess = true;
                            break;
                        }
                        catch (Exception ex)
                        {
                            if (retry < moveRetries - 1)
                            {
                                int waitTime = 500 * (retry + 1); // 递增等待：500ms, 1000ms, 1500ms, 2000ms, 2500ms
                                logger.LogWarning($"移动临时文件失败（第 {retry + 1} 次尝试）: {ex.Message}，等待 {waitTime}ms 后重试...");
                                System.Threading.Thread.Sleep(waitTime);
                                
                                // 尝试强制垃圾回收，释放可能的文件句柄
                                System.GC.Collect();
                                System.GC.WaitForPendingFinalizers();
                                System.GC.Collect();
                            }
                            else
                            {
                                logger.LogError($"移动临时文件失败: {ex.Message}");
                                // 尝试删除临时文件
                                try { File.Delete(tempPath); } catch { }
                                return false;
                            }
                        }
                    }
                    
                    if (!moveSuccess)
                    {
                        logger.LogError("[SaveAs] 移动临时文件失败，已达到最大重试次数");
                        try { File.Delete(tempPath); } catch { }
                        return false;
                    }
                }
                
                // 验证文件是否成功保存
                if (!File.Exists(absolutePath))
                {
                    logger.LogError("保存文件后文件不存在，可能保存失败");
                    return false;
                }

                logger.LogSuccess($"文件已保存: {absolutePath} (大小: {new FileInfo(absolutePath).Length} 字节)");
                return true;
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                logger.LogError($"COM 错误：保存文件失败: {ex.Message} (HRESULT: 0x{ex.ErrorCode:X})");
                return false;
            }
            catch (UnauthorizedAccessException ex)
            {
                logger.LogError($"访问被拒绝，无法保存文件: {ex.Message}");
                return false;
            }
            catch (IOException ex)
            {
                logger.LogError($"IO 错误：保存文件失败: {ex.Message}");
                return false;
            }
            catch (Exception ex)
            {
                logger.LogError($"保存文件失败: {ex.Message}\n堆栈跟踪: {ex.StackTrace}");
                return false;
            }
        }

        /// <summary>
        /// 关闭文档
        /// </summary>
        public void Close()
        {
            var logger = new Logger();
            try
            {
                if (_presentation != null)
                {
                    logger.LogInfo("[Close] 准备关闭演示文稿");
                    _presentation.Close();
                    logger.LogInfo("[Close] _presentation.Close() 调用返回");
                    VstoHelper.ReleaseComObject(_presentation);
                    logger.LogInfo("[Close] COM 对象已释放");
                    _presentation = null;
                }
            }
            catch (Exception ex)
            {
                logger.LogWarning($"关闭演示文稿时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 清理资源
        /// </summary>
        private void Cleanup()
        {
            var logger = new Logger();
            try
            {
                logger.LogInfo("[Cleanup] 开始清理资源");
                
                Close();

                if (_app != null)
                {
                    // ⭐ 策略1：智能实例管理 - 只有在我们创建了实例时才关闭应用程序
                    if (_appCreatedByUs)
                    {
                        // 检查是否还有其他演示文稿打开
                        int remainingPresentations = 0;
                        try
                        {
                            remainingPresentations = _app.Presentations.Count;
                        }
                        catch (Exception ex)
                        {
                            logger.LogWarning($"检查演示文稿数量时出错: {ex.Message}");
                        }
                        
                        if (remainingPresentations == 0)
                        {
                            logger.LogInfo("[Cleanup] 准备关闭 PowerPoint 应用程序（我们创建的实例，且无其他演示文稿）");
                            try
                            {
                                _app.Quit();
                                logger.LogInfo("[Cleanup] _app.Quit() 调用返回");
                            }
                            catch (Exception ex)
                            {
                                logger.LogWarning($"关闭 PowerPoint 应用程序时出错: {ex.Message}");
                            }
                        }
                        else
                        {
                            logger.LogInfo($"[Cleanup] PowerPoint 应用程序仍有 {remainingPresentations} 个演示文稿打开，不关闭应用程序");
                        }
                    }
                    else
                    {
                        logger.LogInfo("[Cleanup] PowerPoint 实例不是我们创建的，不关闭应用程序");
                    }
                    
                    // 释放 COM 对象
                    VstoHelper.ReleaseComObject(_app);
                    logger.LogInfo("[Cleanup] PowerPoint 应用程序 COM 对象已释放");
                    _app = null;
                }

                // 强制垃圾回收以释放 COM 对象
                logger.LogInfo("[Cleanup] 准备强制垃圾回收");
                VstoHelper.ForceGarbageCollection();
                logger.LogInfo("[Cleanup] 垃圾回收完成，资源清理结束");
            }
            catch (Exception ex)
            {
                logger.LogWarning($"清理资源时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 释放资源
        /// </summary>
        public void Dispose()
        {
            if (!_disposed)
            {
                Cleanup();
                _disposed = true;
            }
        }
    }
}

