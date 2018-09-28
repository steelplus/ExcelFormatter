using Microsoft.Extensions.CommandLineUtils;
using ExcelFormatter.src;
using System;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelFormatter
{
    /// <author>
    /// Keisuke Hayase
    /// </author>
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            var application = new CommandLineApplication(throwOnUnexpectedArg: false)
            {
                Name = nameof(ExcelFormatter),
            };

            application.HelpOption("-?|-h|--help");

            var SelectA1 = application.Option("-a1|--selecta1", "指定のシートに対して、「A1セル」を選択した状態にします。", CommandOptionType.NoValue);
            var copyTemplate = application.Option("-te|--template", "指定のシートに対してテンプレートを適用します。", CommandOptionType.NoValue);
            var fixPageSetup = application.Option("-pa|--page", "指定のシートに対してページ設定を適用します。", CommandOptionType.NoValue);
            var fixFont = application.Option("-fo|--font", "指定のシートに対してフォント設定を適用します。", CommandOptionType.NoValue);
            var fixHeader = application.Option("-hd|--header", "指定のシートに対してヘッダ設定を適用します。", CommandOptionType.NoValue);
            var fixFooter = application.Option("-ft|--footer", "指定のシートに対してフッタ設定を適用します。", CommandOptionType.NoValue);
            var removeSampleText = application.Option("-sr|--removesample", "指定のシートに対してSAMPLE文字を取り除きます。", CommandOptionType.NoValue);
            var addSampleText = application.Option("-sa|--addsample", "指定のシートに対してSAMPLE文字を追加します。", CommandOptionType.NoValue);
            var overWriteWriteTime = application.Option("-ts|--timestamp", "対象ファイルのタイムスタンプを上書きします。", CommandOptionType.NoValue);
            var mirror = application.Option("-mr|--mirror", "処理対象ファイルを上書きせず、<output dir>に処理後のファイルをコピーします。", CommandOptionType.SingleValue);
            var subDirectory = application.Option("-sd|--subDirectory", "処理対象について、サブフォルダも検索します。", CommandOptionType.NoValue);
            var createConfig = application.Option("-c|--config", "コンフィグファイルの雛形を作成し、アプリケーションを終了します。", CommandOptionType.NoValue);

            var directoryList = application.Argument("dir", "処理対象ディレクトリ", true);

            // アプリケーション実行内容
            application.OnExecute(() =>
            {
                ConsoleColorWriter.SetToConsole();

                // コンフィグファイルの雛形を作成
                if (createConfig.HasValue())
                {
                    ConfigCreator.Save();
                    return 1;
                }

                if (directoryList.Values.Count < 1)
                {
                    application.ShowHelp();
                    return 1;
                }

                // 実行プロセス
                ExecuteOptions executeOptions = new ExecuteOptions
                {
                    SelectA1 = SelectA1.HasValue(),
                    CopyTemplate = copyTemplate.HasValue(),
                    FixPageSetup = fixPageSetup.HasValue(),
                    FixFont = fixFont.HasValue(),
                    FixHeader = fixHeader.HasValue(),
                    FixFooter = fixFooter.HasValue(),
                    RemoveSampleText = removeSampleText.HasValue(),
                    AddSampleText = addSampleText.HasValue(),
                    OverWriteWriteTime = overWriteWriteTime.HasValue(),
                    SubDirectory = subDirectory.HasValue(),
                    MirrorDirectory = mirror.Value()
                };

                if (!executeOptions.HasValue())
                {
                    LogUtil.LogWarn("実行オプションが指定されていません。");
                    application.ShowHelp();
                    return 1;
                }

                if (!String.IsNullOrEmpty(mirror.Value()) && !System.IO.Directory.Exists(mirror.Value()))
                {
                    LogUtil.LogError("[-mr] オプションの値に存在しないディレクトリ名が含まれています。");
                    return 1;
                }

                using (ExcelFormatterApp excelApp = new ExcelFormatterApp())
                {
                    try
                    {
                        excelApp.Open();

                        if (!ConfigValidator.Validate(excelApp.Config))
                        {
                            return 1;
                        }

                        // 選択したオプションをコンソールに表示
                        Console.WriteLine("----------------------------------------");
                        Console.WriteLine("以下の操作を実行します：");
                        foreach (var option in application.Options)
                        {
                            if (option.HasValue())
                            {
                                Console.WriteLine(option.Description);
                            }
                        }
                        Console.WriteLine("----------------------------------------");

                        FormatExcelFiles(directoryList, executeOptions, excelApp);
                    }
                    catch (ApplicationException e)
                    {
                        LogUtil.LogWarn("アプリケーションの初期化に失敗しました。", e);
                        return 1;
                    }
                    finally
                    {
                        excelApp.Close();
                        Console.WriteLine("----------------------------------------");
                        LogUtil.Log("処理が終了しました。");
                        ConsoleColorWriter.ResetToConsole();
                    }
                }

                return 0;
            });
            try
            {
                application.Execute(args);
            }
            catch (CommandParsingException)
            {
                application.ShowHelp();
            }
        }

        /// <summary>
        /// エクセルファイル群をフォーマットします。
        /// </summary>
        /// <param name="directoryList"></param>
        /// <param name="executeOptions"></param>
        /// <param name="excelApp"></param>
        private static void FormatExcelFiles(CommandArgument directoryList, ExecuteOptions executeOptions, ExcelFormatterApp excelApp)
        {
            foreach (string directory in directoryList.Values)
            {
                // ディレクトリ下のxlsxファイルをすべて取得
                DirectoryInfo di = new DirectoryInfo(directory);
                // ディレクトリ検索オプション
                SearchOption searchOption = executeOptions.SubDirectory ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
                try
                {
                    FileInfo[] files = System.Array.FindAll(di.GetFiles("*", searchOption), SearchFile);

                    Console.WriteLine("----------------------------------------");
                    LogUtil.Log("ディレクトリ[" + directory + "]内のエクセルファイルに対し処理を開始します。");
                    foreach (FileInfo file in files)
                    {

                        // エクセルファイル一つ一つに対して処理を行う
                        using (ExcelFormatter formatter = new ExcelFormatter(file.FullName, excelApp, excelApp.Config, executeOptions, directory))
                        {
                            try
                            {
                                formatter.Open();
                                formatter.Format();
                                formatter.SaveAndClose();
                            }
                            catch (Exception e)
                            {
                                LogUtil.LogWarn("[" + file + "] の処理中にエラーが発生しました。対象への処理を停止します。", e);
                            }
                        }
                    }
                }
                catch (System.IO.DirectoryNotFoundException e)
                {
                    LogUtil.LogError("[" + directory + "]は存在しないディレクトリです。対象への処理を停止します。", e);
                }
            }
        }

        /// <summary>
        /// ディレクトリを受け取りエクセルファイルを検索します。
        /// </summary>
        /// <param name="file"></param>
        /// <returns></returns>
        static bool SearchFile(FileInfo file)
        {
            return System.Text.RegularExpressions.Regex.IsMatch(
                file.Name,
                "\\.(?:xlsx|xls)$",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        }
    }
}
