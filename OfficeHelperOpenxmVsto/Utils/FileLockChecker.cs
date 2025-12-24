using System;
using System.Diagnostics;
using System.IO;

namespace OfficeHelperOpenXml.Utils
{
    /// <summary>
    /// 检查文件被哪个进程占用的工具类
    /// </summary>
    public static class FileLockChecker
    {
        /// <summary>
        /// 检查文件被哪些进程占用
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <returns>占用文件的进程信息</returns>
        public static string CheckFileLock(string filePath)
        {
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            {
                return "文件不存在或路径为空";
            }

            try
            {
                var result = new System.Text.StringBuilder();
                result.AppendLine($"检查文件: {filePath}");
                result.AppendLine("=".PadRight(60, '='));

                // 方法1: 尝试打开文件来检测锁定
                result.AppendLine("\n[方法1] 尝试打开文件检测:");
                try
                {
                    using (var fs = File.Open(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                    {
                        result.AppendLine("  - 文件未被锁定，可以正常访问");
                    }
                }
                catch (IOException ioEx)
                {
                    result.AppendLine($"  - 文件被锁定: {ioEx.Message}");
                }
                catch (UnauthorizedAccessException uaEx)
                {
                    result.AppendLine($"  - 访问被拒绝: {uaEx.Message}");
                }

                // 方法2: 检查所有 PowerPoint 进程
                result.AppendLine("\n[方法2] 检查所有 PowerPoint 进程:");
                var pptProcesses = Process.GetProcessesByName("POWERPNT");
                if (pptProcesses.Length > 0)
                {
                    result.AppendLine($"  找到 {pptProcesses.Length} 个 PowerPoint 进程:");
                    foreach (var proc in pptProcesses)
                    {
                        try
                        {
                            result.AppendLine($"  - PID: {proc.Id}, 名称: {proc.ProcessName}, 启动时间: {proc.StartTime:yyyy-MM-dd HH:mm:ss}");
                            try
                            {
                                result.AppendLine($"    主模块: {proc.MainModule?.FileName ?? "N/A"}");
                            }
                            catch
                            {
                                result.AppendLine($"    主模块: 无法访问");
                            }
                        }
                        catch
                        {
                            result.AppendLine($"  - PID: {proc.Id}, 名称: {proc.ProcessName}");
                        }
                    }
                }
                else
                {
                    result.AppendLine("  - 未找到 PowerPoint 进程");
                }

                // 方法3: 检查所有可能占用 Office 文件的进程
                result.AppendLine("\n[方法3] 检查其他可能占用文件的进程:");
                string[] officeProcessNames = { "WINWORD", "EXCEL", "OUTLOOK", "MSACCESS" };
                bool foundOfficeProcess = false;
                foreach (var procName in officeProcessNames)
                {
                    var processes = Process.GetProcessesByName(procName);
                    if (processes.Length > 0)
                    {
                        foundOfficeProcess = true;
                        result.AppendLine($"  找到 {processes.Length} 个 {procName} 进程");
                        foreach (var proc in processes)
                        {
                            try
                            {
                                result.AppendLine($"    - PID: {proc.Id}, 启动时间: {proc.StartTime:yyyy-MM-dd HH:mm:ss}");
                            }
                            catch
                            {
                                result.AppendLine($"    - PID: {proc.Id}");
                            }
                        }
                    }
                }
                if (!foundOfficeProcess)
                {
                    result.AppendLine("  - 未找到其他 Office 进程");
                }

                // 方法4: 尝试使用 Handle.exe (Sysinternals) 如果可用
                result.AppendLine("\n[方法4] 尝试使用 Handle.exe (Sysinternals):");
                try
                {
                    var handleInfo = GetFileHandlesUsingHandle(filePath);
                    if (!string.IsNullOrEmpty(handleInfo))
                    {
                        result.AppendLine(handleInfo);
                    }
                    else
                    {
                        result.AppendLine("  - Handle.exe 不可用（需要从 Sysinternals 下载）");
                    }
                }
                catch (Exception ex)
                {
                    result.AppendLine($"  - Handle.exe 执行失败: {ex.Message}");
                }

                result.AppendLine("\n提示: 如果文件被锁定，可以：");
                result.AppendLine("  1. 关闭所有 PowerPoint 窗口");
                result.AppendLine("  2. 在任务管理器中结束 POWERPNT.exe 进程");
                result.AppendLine("  3. 使用资源监视器查找占用文件的进程");

                return result.ToString();
            }
            catch (Exception ex)
            {
                return $"检查文件锁定时出错: {ex.Message}\n{ex.StackTrace}";
            }
        }

        /// <summary>
        /// 使用 Handle.exe (Sysinternals) 查询文件句柄
        /// </summary>
        private static string GetFileHandlesUsingHandle(string filePath)
        {
            try
            {
                // 尝试查找 Handle.exe
                string handlePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Sysinternals", "handle.exe");
                if (!File.Exists(handlePath))
                {
                    handlePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "Sysinternals", "handle.exe");
                }

                if (!File.Exists(handlePath))
                {
                    return null; // Handle.exe 不可用
                }

                var startInfo = new ProcessStartInfo
                {
                    FileName = handlePath,
                    Arguments = $"\"{filePath}\"",
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true
                };

                using (var process = Process.Start(startInfo))
                {
                    if (process != null)
                    {
                        string output = process.StandardOutput.ReadToEnd();
                        string error = process.StandardError.ReadToEnd();
                        process.WaitForExit();
                        
                        if (!string.IsNullOrEmpty(output))
                        {
                            return output;
                        }
                        if (!string.IsNullOrEmpty(error))
                        {
                            return $"错误: {error}";
                        }
                    }
                }
            }
            catch
            {
                // Handle.exe 不可用或执行失败
            }

            return null;
        }
    }
}

