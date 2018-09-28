using System;

namespace ExcelFormatter
{
    public class LogUtil
    {
        /// <summary>
        /// 通常のログと必要に応じてスタックトレースを表示します。
        /// </summary>
        /// <param name="message"></param>
        public static void Log(string message, Exception e = null)
        {
            WriteLog("ログ : " + message);
            if (e != null)
            {
                WriteLog("詳細 : ");
                WriteLog(e.ToString());
            }
        }

        /// <summary>
        /// 警告ログと必要に応じてスタックトレースを表示します。
        /// </summary>
        /// <param name="message"></param>
        /// <param name="e"></param>
        public static void LogWarn(string message, Exception e = null)
        {
            ConsoleColorWriter.SetToConsole(ConsoleColor.Yellow);
            WriteLog("警告 : " + message);
            if(e != null)
            {
                WriteLog("詳細 : ");
                WriteLog(e.ToString());
            }
            ConsoleColorWriter.SetToConsole();
        }

        /// <summary>
        /// エラーログと必要に応じてスタックトレースを表示します。
        /// </summary>
        /// <param name="message"></param>
        /// <param name="e"></param>
        public static void LogError(string message, Exception e = null)
        {
            ConsoleColorWriter.SetToConsole(ConsoleColor.Red);
            WriteError("エラー : " + message);
            if (e != null)
            {
                WriteError("詳細 : ");
                WriteError(e.ToString());
            }
            ConsoleColorWriter.SetToConsole();
        }

        /// <summary>
        /// 標準コンソールにタイムスタンプ付のメッセージを出力します。
        /// </summary>
        /// <param name="message"></param>
        private static void WriteLog(string message)
        {
            DateTimeOffset now = DateTimeOffset.Now;
            Console.WriteLine(String.Format("{0}: {1}", now.ToString("yyyy-MM-dd HH:mm:ss.fff"), message));
        }

        /// <summary>
        /// 標準エラーにタイムスタンプ付のメッセージを出力します。
        /// </summary>
        /// <param name="message"></param>
        private static void WriteError(string message)
        {
            DateTimeOffset now = DateTimeOffset.Now;
            Console.Error.WriteLine(String.Format("{0}: {1}", now.ToString("yyyy-MM-dd HH:mm:ss.fff"), message));
        }
    }
}
