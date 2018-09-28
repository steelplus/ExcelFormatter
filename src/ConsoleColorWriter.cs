using System;
using System.IO;
using System.Text;

namespace ExcelFormatter
{
    /// <summary>
    /// コンソール出力を色分けして表示するクラスです。
    /// </summary>
    public class ConsoleColorWriter : TextWriter
    {
        private ConsoleColor foregroundColor;
        private ConsoleColor backgroudnColor;
        private TextWriter originalConsoleStream;
        private static ConsoleColor originalForegroundColor = default(ConsoleColor);
        private static ConsoleColor originalBackgroundColor = default(ConsoleColor);

        /// <summary>
        /// コンストラクタ 
        /// </summary>
        /// <param name="consoleTextWriter"></param>
        /// <param name="foregroundColor"></param>
        /// <param name="backgroudnColor"></param>
        public ConsoleColorWriter(TextWriter consoleTextWriter, ConsoleColor foregroundColor, ConsoleColor backgroudnColor)
        {
            originalConsoleStream = consoleTextWriter;
            this.foregroundColor = foregroundColor;
            this.backgroudnColor = backgroudnColor;
        }

        /// <summary>
        /// メッセージを1行に出力します。
        /// </summary>
        /// <param name="value"></param>
        public override void WriteLine(string value)
        {
            ConsoleColor originalForegroundColor = Console.ForegroundColor;
            ConsoleColor originalBackgroundColor = Console.BackgroundColor;
            Console.ForegroundColor = foregroundColor;
            Console.BackgroundColor = backgroudnColor;
            
            originalConsoleStream.WriteLine(value);

            Console.ForegroundColor = originalForegroundColor;
            Console.BackgroundColor = originalBackgroundColor;
        }

        /// <summary>
        /// エンコードを取得します。
        /// </summary>
        public override Encoding Encoding
        {
            get { return Encoding.Default; }
        }

        /// <summary>
        /// 標準出力を指定した色に設定します。
        /// </summary>
        /// <param name="foregroundColor"></param>
        /// <param name="backgroundColor"></param>
        public static void SetToConsole(ConsoleColor foregroundColor = ConsoleColor.White, ConsoleColor backgroundColor = ConsoleColor.Black)
        {
            if(originalForegroundColor == default(ConsoleColor))
            {
                originalForegroundColor = Console.ForegroundColor;
            }
            if(originalBackgroundColor == default(ConsoleColor))
            {
                originalBackgroundColor = Console.BackgroundColor;
            }
            Console.ForegroundColor = foregroundColor;
            Console.BackgroundColor = backgroundColor;
        }

        public static void ResetToConsole()
        {
            Console.ForegroundColor = originalForegroundColor;
            Console.BackgroundColor = originalBackgroundColor;
        }

        /// <summary>
        /// 標準エラーを指定した色に設定します。
        /// </summary>
        /// <param name="foregroundColor"></param>
        /// <param name="backgroundColor"></param>
        public static void SetToConsoleError(ConsoleColor foregroundColor = ConsoleColor.Red, ConsoleColor backgroundColor = ConsoleColor.Black)
        {
            Console.SetError(new ConsoleColorWriter(Console.Error, foregroundColor, backgroundColor));
        }
    }
}
