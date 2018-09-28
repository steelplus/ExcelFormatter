using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
namespace ExcelFormatter.src
{

    /// <summary>
    /// Excel.Applicationのラッパークラスです。
    /// </summary>
    class ExcelFormatterApp : IDisposable
    {
        public Excel.Application Application { get; set; }
        public Excel.Workbook Template { get; set; }
        public Config Config { get; set; }

        public ExcelFormatterApp()
        {
        }

        public void Open()
        {
            this.Application = new Excel.Application
            {
                Visible = false,
                DisplayAlerts = false,
                AskToUpdateLinks = false
            };
            try
            {
                // テンプレートワークブックのオープン
                this.Template = this.Application.Workbooks.Open(
                    Filename: Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + @"\template.xlsx",
                    UpdateLinks: Excel.XlUpdateLinks.xlUpdateLinksNever,
                    ReadOnly: true,
                    IgnoreReadOnlyRecommended: true
                    );
            }
            catch (Exception)
            {
                throw new ApplicationException("テンプレートファイルを開けませんでした。テンプレートファイルに問題があるか、配置されていません。\r\n実行ファイルと同じディレクトリにtemplate.xlsxを配置して下さい。");
            }

            try
            {
                // コンフィグファイルの読み込み
                this.Config = ConfigSerializer.Deserialize(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + @"\config.json");
            }
            catch (Exception)
            {
                throw new ApplicationException("設定ファイルを開けませんでした。設定ファイルが正しくないか、配置されていません。\r\n実行ファイルと同じディレクトリにconfig.jsonを配置して下さい。");
            }
        }

        /// <summary>
        /// ExcelオブジェクトをCloseします。
        /// </summary>
        public void Close()
        {
            this.Template.Close(false);
            this.Application.Quit();
        }

        /// <summary>
        /// 終端処理。オブジェクトを開放します。
        /// </summary>
        public void Dispose()
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(this.Template);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(this.Application);
        }
    }
}
