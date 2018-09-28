using ExcelFormatter.src;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelFormatter
{
    /// <summary>
    /// 設定に従いExcelフォーマットを行うクラスです。
    /// </summary>
    /// <author>
    /// Keisuke Hayase
    /// </author>
    class ExcelFormatter : IDisposable
    {
        private readonly Excel.Application ExcelApp;

        // 処理対象ブック
        private Excel.Workbook Workbook;

        // テンプレートブック
        private readonly Excel.Workbook TemplateWorkBook;

        // 実行オプション
        private readonly ExecuteOptions Options;

        // 設定
        private Config Config;

        // シートのファイルパス
        private readonly String filePath;

        // ファイルのタイムスタンプ
        private readonly DateTime CreationTime;
        private readonly DateTime LastWriteTime;
        private readonly DateTime LastAccessTime;

        // 読み込みディレクトリ
        private readonly String Directory;

        /// <summary>
        /// 開始処理
        /// </summary>
        /// <param name="filePath"></param>
        public ExcelFormatter(string filePath, ExcelFormatterApp eApp, Config config, ExecuteOptions options, string directory)
        {
            // アプリケーションの代入
            this.ExcelApp = eApp.Application;
            // テンプレートワークブックの読み込み
            this.TemplateWorkBook = eApp.Template;
            // コンフィグファイルの読み込み
            this.Config = config;
            // 実行オプションの読み込み
            this.Options = options;

            // 更新時刻の記録
            if (!this.Options.OverWriteWriteTime)
            {
                this.CreationTime = File.GetCreationTime(filePath);
                this.LastAccessTime = File.GetLastAccessTime(filePath);
                this.LastWriteTime = File.GetLastWriteTime(filePath);
            }
            this.filePath = filePath;
            this.Directory = directory;
        }

        public void Open()
        {
            // ブックのオープン
            this.Workbook = this.ExcelApp.Workbooks.Open(
                Filename: System.IO.Path.GetFullPath(filePath),
                UpdateLinks: Excel.XlUpdateLinks.xlUpdateLinksAlways,
                ReadOnly: false,
                IgnoreReadOnlyRecommended: true
                );
        }

        /// <summary>
        /// エクセルファイルをフォーマットします。
        /// </summary>
        public void Format()
        {
            // ログ
            Console.WriteLine("----------------------------------------");
            LogUtil.Log("[" + Workbook.FullName + "]に対する処理を開始しました。");

            // シート一覧を取得
            for (int i = 1; i <= this.Workbook.Sheets.Count; i++)
            {
                Excel.Worksheet worksheet = this.Workbook.Sheets[i];
                try
                {
                    if (worksheet.Visible == Excel.XlSheetVisibility.xlSheetVisible)
                    {
                        // フォーマット実行
                        if (this.Options.CopyTemplate)
                        {
                            foreach (dynamic template in this.Config.Template)
                            {
                                if (IsApplySheet(worksheet.Name, template))
                                {
                                    CopyTemplate(worksheet, template);
                                    break;
                                }
                            }
                        }
                        if (this.Options.FixPageSetup)
                        {
                            foreach (dynamic pageSetup in this.Config.PageSetup)
                            {
                                if (IsApplySheet(worksheet.Name, pageSetup))
                                {
                                    FixPageSetup(worksheet, pageSetup);
                                    break;
                                }
                            }
                        }
                        if (this.Options.FixFont)
                        {
                            foreach (dynamic font in this.Config.Font)
                            {
                                if (IsApplySheet(worksheet.Name, font))
                                {
                                    FixFont(worksheet, font);
                                    break;
                                }
                            }
                        }
                        if (this.Options.FixHeader)
                        {
                            foreach (dynamic header in this.Config.Header)
                            {
                                if (IsApplySheet(worksheet.Name, header))
                                {
                                    FixHeader(worksheet, header);
                                    break;
                                }
                            }
                        }
                        if (this.Options.FixFooter)
                        {
                            foreach (dynamic footer in this.Config.Footer)
                            {
                                if (IsApplySheet(worksheet.Name, footer))
                                {
                                    FixFooter(worksheet, footer);
                                    break;
                                }
                            }
                        }
                        if (this.Options.RemoveSampleText)
                        {
                            foreach (dynamic removeSample in this.Config.RemoveSample)
                            {
                                if (IsApplySheet(worksheet.Name, removeSample))
                                {
                                    RemoveSampleText(worksheet, removeSample);
                                    break;
                                }
                            }
                        }
                        if (this.Options.AddSampleText)
                        {
                            foreach (dynamic addSample in this.Config.AddSample)
                            {
                                if (IsApplySheet(worksheet.Name, addSample))
                                {
                                    AddSampleText(worksheet, addSample);
                                    break;
                                }
                            }
                        }
                        if (this.Options.SelectA1)
                        {
                            if (IsApplySheet(worksheet.Name, Config.SelectA1))
                            {
                                // この処理は最後に
                                SelectA1(worksheet);
                            }
                        }
                    }
                }
                finally
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                }
            }
        }

        /// <summary>
        /// シートの「SAMPLE」文字をすべて削除します。
        /// </summary>
        /// <param name="worksheet"></param>
        private void RemoveSampleText(Excel.Worksheet worksheet, RemoveSample removeSample)
        {
            foreach (dynamic shape in worksheet.Shapes)
            {
                try
                {
                    if (shape.TextFrame2.TextRange.Count == 1 && shape.TextFrame2.TextRange.Text == removeSample.Text)
                    {
                        // 削除
                        shape.Delete();
                    }
                }
                catch (Exception e)
                {
                    // コメント等、TextFrame2やTextRangeを実装していないクラスに対してはExceptionをスルーします
                    if (!(e is System.NotImplementedException || e is System.ArgumentException))
                    {
                        throw e;
                    }
                }
            }
        }

        /// <summary>
        /// シートに「SAMPLE」文字を挿入します。
        /// </summary>
        /// <param name="worksheet"></param>
        private void AddSampleText(Excel.Worksheet worksheet, AddSample addSample)
        {
            // テンプレートから図形を取得
            Excel.Worksheet sampleSheet = this.TemplateWorkBook.Sheets[addSample.SampleArtSheet];
            System.Collections.IEnumerator index = sampleSheet.Shapes.GetEnumerator();
            Excel.Shape shape;
            while (index.MoveNext())
            {
                dynamic sampleShape = index.Current;
                if (sampleShape.TextFrame2.TextRange.Count == 1)
                {
                    // クリップボードの内容をクリア
                    Clipboard.Clear();

                    shape = sampleShape;
                    shape.Copy();
                    // ワークシートの所定の位置を選択しコピー
                    worksheet.Activate();

                    int blankCount = 0;
                    int i = addSample.OffsetT;
                    for (; i < addSample.EndOfRow; i++)
                    {
                        // 行が空でないかをチェック
                        bool blankFlag = true;
                        for (int j = 1; j < addSample.EndOfColumn; j++)
                        {
                            try
                            {
                                if (worksheet.Range[GetColumnLetter(j) + i].Value != null)
                                {
                                    blankFlag = false;
                                }
                            }
                            catch (Exception e)
                            {
                                Console.WriteLine(e);
                            }
                        }
                        if (blankFlag)
                        {
                            blankCount += 1;
                            if (blankCount > addSample.EndBlankLine)
                            {
                                break;
                            }
                        }

                        // サンプル画像を貼り付ける
                        if (i == addSample.OffsetT || (i - addSample.OffsetT) % addSample.Interval == 0)
                        {
                            worksheet.Cells[i, addSample.OffsetL].Select();
                            worksheet.Paste();
                        }
                    }
                }
            }
        }

        /// <summary>
        /// シート全体のフォントを整えます。
        /// </summary>
        /// <param name="worksheet"></param>
        private void FixFont(Excel.Worksheet worksheet, Font font)
        {
            // 全セルのフォントを統一
            Excel.Range range = worksheet.UsedRange;
            range.Font.Name = font.FontName;

            // 全図形のテキストのフォントを統一
            foreach (dynamic shape in worksheet.Shapes)
            {
                try
                {
                    if (shape is Excel.Shape && shape.TextFrame2.TextRange != null)
                    {
                        shape.TextFrame2.TextRange.Font.Name = font.FontName;
                        shape.TextFrame2.TextRange.Font.NameFarEast = font.FontName;
                    }
                }
                catch (Exception e)
                {
                    // コメント等、TextFrame2やTextRangeを実装していないクラスに対してはExceptionをスルーします
                    if (!(e is System.NotImplementedException || e is System.ArgumentException))
                    {
                        throw e;
                    }
                }
            }

            // 全コメントのフォントを統一
            foreach (Excel.Comment comment in worksheet.Comments)
            {
                comment.Shape.TextFrame.Characters().Font.Name = font.FontName;
            }
        }


        /// <summary>
        /// シートの印刷設定を整えます。
        /// </summary>
        /// <param name="worksheet"></param>
        private void FixPageSetup(Excel.Worksheet worksheet, PageSetup pageSetup)
        {
            // 印刷設定
            Dictionary<string, PaperSize> paperSizes = PaperSizesCreator.paperSizes;

            // 印刷設定
            worksheet.PageSetup.TopMargin = ExcelApp.CentimetersToPoints(pageSetup.TopMargin);
            worksheet.PageSetup.BottomMargin = ExcelApp.CentimetersToPoints(pageSetup.BottmonMargin);
            worksheet.PageSetup.LeftMargin = ExcelApp.CentimetersToPoints(pageSetup.LeftMargin);
            worksheet.PageSetup.RightMargin = ExcelApp.CentimetersToPoints(pageSetup.RightMargin);
            worksheet.PageSetup.HeaderMargin = ExcelApp.CentimetersToPoints(pageSetup.HeaderMargin);
            worksheet.PageSetup.FooterMargin = ExcelApp.CentimetersToPoints(pageSetup.FooterMargin);
            worksheet.PageSetup.Zoom = pageSetup.Zoom;
            worksheet.PageSetup.PaperSize = (Excel.XlPaperSize)PaperSizesCreator.paperSizes[pageSetup.PaperSize].Value;
            worksheet.PageSetup.Orientation = (Excel.XlPageOrientation)PageOrientationsCreator.pageOrientations[pageSetup.Orientation].Value;
        }

        /// <summary>
        /// シートのヘッダ、フッタを整えます。
        /// </summary>
        /// <param name="worksheet"></param>
        private void FixHeader(Excel.Worksheet worksheet, Header header)
        {
            // ヘッダ・フッタ部
            worksheet.PageSetup.CenterHeader = header.CenterHeader;
            worksheet.PageSetup.LeftHeader = header.LeftHeader;
            worksheet.PageSetup.RightHeader = header.RightHeader;
        }

        /// <summary>
        /// シートのヘッダ、フッタを整えます。
        /// </summary>
        /// <param name="worksheet"></param>
        private void FixFooter(Excel.Worksheet worksheet, Footer footer)
        {
            // ヘッダ・フッタ部
            worksheet.PageSetup.CenterFooter = footer.CenterFooter;
            worksheet.PageSetup.LeftFooter = footer.LeftFooter;
            worksheet.PageSetup.RightFooter = footer.RightFooter;
        }

        /// <summary>
        /// ヘッダファイルをテンプレートからコピーします
        /// </summary>
        private void CopyTemplate(Excel.Worksheet worksheet, Template template)
        {
            // テンプレートファイルの指定のシートを読み込み
            Excel.Worksheet templateSheet = this.TemplateWorkBook.Sheets[template.TemplateSheetNum];

            try
            {
                // テンプレートからセルのコピー
                templateSheet.Range[template.TemplateCellFrom, template.TemplateCellEnd].Copy(worksheet.Range[template.TemplateCellFrom, template.TemplateCellEnd]);
                // コンフィグファイルから値のコピー
                foreach (TemplateVal templateVal in template.TemplateVals)
                {
                    worksheet.Range[templateVal.Cell].Value2 = templateVal.Value;
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(templateSheet);
            }
        }

        /// <summary>
        /// 指定された正規表現から、シート名がマッチするかどうかを返します
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        private bool IsApplySheet(string sheetName, dynamic config)
        {
            // 除外正規表現の探索
            if (!String.IsNullOrEmpty(config.ExcludeTemplateSheetRegex) && Regex.IsMatch(sheetName, config.ExcludeTemplateSheetRegex))
            {
                return false;
            }

            // 適用正規表現の探索
            if (String.IsNullOrEmpty(config.ApplyTemplateSheetRegex) || Regex.IsMatch(sheetName, config.ApplyTemplateSheetRegex))
            {
                return true;
            }

            return false;
        }

        /// <summary>
        /// 選択をA1セルに固定します
        /// </summary>
        /// <param name="worksheet"></param>
        private void SelectA1(dynamic worksheet)
        {
            worksheet.Activate();
            worksheet.Range["A1"].Select();
            worksheet.Range["A1"].Activate();
        }

        private static string GetColumnLetter(int column)
        {
            if (column < 1) return String.Empty;
            return GetColumnLetter((column - 1) / 26) + (char)('A' + (column - 1) % 26);
        }

        /// <summary>
        /// 保存するファイルパスを取得します。
        /// </summary>
        /// <returns></returns>
        private string GetSaveFilePath()
        {
            if (String.IsNullOrEmpty(Options.MirrorDirectory))
            {
                return Path.GetFullPath(this.filePath);
            }
            string fullPath = Path.GetFullPath(Options.MirrorDirectory) + this.filePath.Substring(this.Directory.Length);
            if (!System.IO.Directory.Exists(Path.GetDirectoryName(fullPath)))
            {
                System.IO.Directory.CreateDirectory(Path.GetDirectoryName(fullPath));
            }
            return fullPath;
        }

        /// <summary>
        /// ブックの内容を保存し、Closeします。
        /// </summary>
        public void SaveAndClose()
        {
            // 表紙シート（表示シート1枚目をアクティブにする）
            foreach (Excel.Worksheet sheet in this.Workbook.Sheets)
            {
                if (sheet.Visible == Excel.XlSheetVisibility.xlSheetVisible)
                {
                    sheet.Activate();
                    break;
                }
            }

            // 保存
            string writePath = GetSaveFilePath();
            this.Workbook.SaveAs(writePath);
            this.Workbook.Close();

            // 更新日時の上書き
            if (!this.Options.OverWriteWriteTime)
            {
                File.SetCreationTime(writePath, this.CreationTime);
                File.SetLastAccessTime(writePath, this.LastAccessTime);
                File.SetLastWriteTime(writePath, this.LastWriteTime);
            }
        }

        /// <summary>
        /// 終端処理
        /// </summary>
        public void Dispose()
        {
            if (this.Workbook != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(this.Workbook);
            }
        }
    }
}
