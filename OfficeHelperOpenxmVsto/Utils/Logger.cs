using System;

namespace OfficeHelperOpenXml.Utils
{
    /// <summary>
    /// Simple logger for general purpose logging
    /// </summary>
    public class Logger
    {
        /// <summary>
        /// Logs an informational message
        /// </summary>
        /// <param name="message">The message to log</param>
        public void LogInfo(string message)
        {
            Console.WriteLine($"[INFO] {message}");
        }

        /// <summary>
        /// Logs a success message
        /// </summary>
        /// <param name="message">The message to log</param>
        public void LogSuccess(string message)
        {
            Console.WriteLine($"[SUCCESS] {message}");
        }

        /// <summary>
        /// Logs a warning message
        /// </summary>
        /// <param name="message">The message to log</param>
        public void LogWarning(string message)
        {
            Console.WriteLine($"[WARNING] {message}");
        }

        /// <summary>
        /// Logs an error message
        /// </summary>
        /// <param name="message">The message to log</param>
        public void LogError(string message)
        {
            Console.WriteLine($"[ERROR] {message}");
        }

        /// <summary>
        /// Logs progress information for tracking progress
        /// </summary>
        /// <param name="operation">The operation being performed</param>
        /// <param name="current">The current item number</param>
        /// <param name="total">The total number of items</param>
        public void LogProgress(string operation, int current, int total)
        {
            Console.WriteLine($"[PROGRESS] {operation}: {current}/{total}");
        }
    }
}

