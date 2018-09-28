using ExcelFormatter.src;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;

namespace ExcelFormatter
{

    [DataContract]
    public class SelectA1
    {

        [DataMember(Name = "applyTemplateSheetRegex")]
        public string ApplyTemplateSheetRegex { get; set; }

        [DataMember(Name = "excludeTemplateSheetRegex")]
        public string ExcludeTemplateSheetRegex { get; set; }
    }

    [DataContract]
    public class RemoveSample
    {

        [DataMember(Name = "applyTemplateSheetRegex")]
        public string ApplyTemplateSheetRegex { get; set; }

        [DataMember(Name = "excludeTemplateSheetRegex")]
        public string ExcludeTemplateSheetRegex { get; set; }

        [DataMember(Name = "text")]
        public string Text { get; set; }
    }

    [DataContract]
    public class AddSample
    {

        [DataMember(Name = "applyTemplateSheetRegex")]
        public string ApplyTemplateSheetRegex { get; set; }

        [DataMember(Name = "excludeTemplateSheetRegex")]
        public string ExcludeTemplateSheetRegex { get; set; }

        [DataMember(Name = "sampleArtSheet")]
        public int SampleArtSheet { get; set; }

        [DataMember(Name = "offsetL")]
        public int OffsetL { get; set; }

        [DataMember(Name = "offsetT")]
        public int OffsetT { get; set; }

        [DataMember(Name = "interval")]
        public int Interval { get; set; }

        [DataMember(Name = "endOfColumn")]
        public int EndOfColumn { get; set; }

        [DataMember(Name = "endOfRow")]
        public int EndOfRow { get; set; }

        [DataMember(Name = "endBlankLine")]
        public int EndBlankLine { get; set; }
    }

    [DataContract]
    public class Font
    {

        [DataMember(Name = "applyTemplateSheetRegex")]
        public string ApplyTemplateSheetRegex { get; set; }

        [DataMember(Name = "excludeTemplateSheetRegex")]
        public string ExcludeTemplateSheetRegex { get; set; }

        [DataMember(Name = "fontName")]
        public string FontName { get; set; }
    }

    [DataContract]
    public class TemplateVal
    {

        [DataMember(Name = "cell")]
        public string Cell { get; set; }

        [DataMember(Name = "value")]
        public string Value { get; set; }
    }

    [DataContract]
    public class Template
    {

        [DataMember(Name = "applyTemplateSheetRegex")]
        public string ApplyTemplateSheetRegex { get; set; }

        [DataMember(Name = "excludeTemplateSheetRegex")]
        public string ExcludeTemplateSheetRegex { get; set; }

        [DataMember(Name = "templateSheetNum")]
        public int TemplateSheetNum { get; set; }

        [DataMember(Name = "templateCellFrom")]
        public string TemplateCellFrom { get; set; }

        [DataMember(Name = "templateCellEnd")]
        public string TemplateCellEnd { get; set; }

        [DataMember(Name = "templateVals")]
        public IList<TemplateVal> TemplateVals { get; set; }
    }

    [DataContract]
    public class PageSetup
    {

        [DataMember(Name = "applyTemplateSheetRegex")]
        public string ApplyTemplateSheetRegex { get; set; }

        [DataMember(Name = "excludeTemplateSheetRegex")]
        public string ExcludeTemplateSheetRegex { get; set; }

        [DataMember(Name = "topMargin")]
        public double TopMargin { get; set; }

        [DataMember(Name = "bottmonMargin")]
        public double BottmonMargin { get; set; }

        [DataMember(Name = "leftMargin")]
        public double LeftMargin { get; set; }

        [DataMember(Name = "rightMargin")]
        public double RightMargin { get; set; }

        [DataMember(Name = "headerMargin")]
        public double HeaderMargin { get; set; }

        [DataMember(Name = "footerMargin")]
        public double FooterMargin { get; set; }

        [DataMember(Name = "zoom")]
        public int Zoom { get; set; }

        [DataMember(Name = "paperSize")]
        public string PaperSize { get; set; }

        [DataMember(Name = "orientation")]
        public string Orientation { get; set; }
    }

    [DataContract]
    public class Header
    {

        [DataMember(Name = "applyTemplateSheetRegex")]
        public string ApplyTemplateSheetRegex { get; set; }

        [DataMember(Name = "excludeTemplateSheetRegex")]
        public string ExcludeTemplateSheetRegex { get; set; }

        [DataMember(Name = "centerHeader")]
        public string CenterHeader { get; set; }

        [DataMember(Name = "leftHeader")]
        public string LeftHeader { get; set; }

        [DataMember(Name = "rightHeader")]
        public string RightHeader { get; set; }
    }

    [DataContract]
    public class Footer
    {

        [DataMember(Name = "applyTemplateSheetRegex")]
        public string ApplyTemplateSheetRegex { get; set; }

        [DataMember(Name = "excludeTemplateSheetRegex")]
        public string ExcludeTemplateSheetRegex { get; set; }

        [DataMember(Name = "centerFooter")]
        public string CenterFooter { get; set; }

        [DataMember(Name = "leftFooter")]
        public string LeftFooter { get; set; }

        [DataMember(Name = "rightFooter")]
        public string RightFooter { get; set; }
    }

    [DataContract]
    public class Config
    {

        [DataMember(Name = "selectA1")]
        public SelectA1 SelectA1 { get; set; }

        [DataMember(Name = "removeSample")]
        public IList<RemoveSample> RemoveSample { get; set; }

        [DataMember(Name = "addSample")]
        public IList<AddSample> AddSample { get; set; }

        [DataMember(Name = "font")]
        public IList<Font> Font { get; set; }

        [DataMember(Name = "template")]
        public IList<Template> Template { get; set; }

        [DataMember(Name = "pageSetup")]
        public IList<PageSetup> PageSetup { get; set; }

        [DataMember(Name = "header")]
        public IList<Header> Header { get; set; }

        [DataMember(Name = "footer")]
        public IList<Footer> Footer { get; set; }
    }

    public class ConfigSerializer
    {
        public static Config Deserialize(string filePath)
        {
            DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(Config));
            using (FileStream fs = new FileStream(filePath, FileMode.Open))
            {
                fs.Position = (fs.ReadByte() == 0xef) ? 3 : 0;
                return (Config)serializer.ReadObject(fs);
            };
        }
    }

    public class ConfigCreator
    {
        /// <summary>
        /// コンフィグファイルの雛形を作成します
        /// </summary>
        public static void Save()
        {
            string configFileName = System.AppDomain.CurrentDomain.BaseDirectory.TrimEnd('\\') + "\\config.json";

            // 既にコンフィグファイルが存在するかどうかを確認
            if (System.IO.File.Exists(configFileName))
            {
                LogUtil.Log("既にコンフィグファイルが存在するため、コンフィグファイルの作成を中止します。");
                return;
            }

            Config config = CreateConfig();

            DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(Config));
            using (FileStream fs = new FileStream(configFileName, FileMode.CreateNew))
            {
                serializer.WriteObject(fs, config);
            }

            LogUtil.Log("コンフィグファイルを作成しました。");
            LogUtil.Log("出力先ディレクトリ：[" + configFileName + "]");
        }

        private static Config CreateConfig()
        {
            SelectA1 selectA1 = new SelectA1
            {
                ApplyTemplateSheetRegex = ".*",
                ExcludeTemplateSheetRegex = ""
            };

            RemoveSample removeSample = new RemoveSample
            {
                ApplyTemplateSheetRegex = ".*",
                ExcludeTemplateSheetRegex = "",
                Text = "SAMPLE"
            };

            Font font = new Font
            {
                ApplyTemplateSheetRegex = ".*",
                ExcludeTemplateSheetRegex = "",
                FontName = "ＭＳ 明朝"
            };

            PageSetup pageSetup = new PageSetup
            {
                ApplyTemplateSheetRegex = ".*",
                ExcludeTemplateSheetRegex = "",
                TopMargin = 1,
                BottmonMargin = 1,
                LeftMargin = 1,
                RightMargin = 1,
                HeaderMargin = 0.5,
                FooterMargin = 0.5,
                Zoom = 100,
                PaperSize = "xlPaperA4",
                Orientation = "xlLandscape"
            };

            Header header = new Header
            {
                ApplyTemplateSheetRegex = ".*",
                ExcludeTemplateSheetRegex = "",
                CenterHeader = "",
                LeftHeader = "",
                RightHeader = ""
            };

            Footer footer = new Footer
            {
                ApplyTemplateSheetRegex = ".*",
                ExcludeTemplateSheetRegex = "",
                CenterFooter = "",
                LeftFooter = "",
                RightFooter = ""
            };

            Config config = new Config
            {
                SelectA1 = selectA1,
                RemoveSample = new List<RemoveSample> { removeSample },
                AddSample = new List<AddSample>(),
                Template = new List<Template>(),
                Font = new List<Font> { font },
                PageSetup = new List<PageSetup> { pageSetup },
                Header = new List<Header> { header },
                Footer = new List<Footer> { footer }
            };

            return config;
        }
    }

    /// <summary>
    /// configに正常な値が入っているかどうかを精査するクラスです。
    /// </summary>
    public class ConfigValidator
    {
        public static bool Validate(Config config)
        {
            List<string> errors = new List<string>();

            // PaperSize
            foreach (PageSetup pageSetup in config.PageSetup)
            {
                if (!PaperSizesCreator.paperSizes.ContainsKey(pageSetup.PaperSize))
                {
                    errors.Add("paperSizeには規定の値を入力して下さい。");
                    foreach (string key in PaperSizesCreator.paperSizes.Keys)
                    {
                        errors.Add("[" + key + "] : ページ設定 - " + PaperSizesCreator.paperSizes[key].Description);
                    }
                    break;
                }
            }

            // Orientation
            foreach (PageSetup pageSetup in config.PageSetup)
            {
                if (!PageOrientationsCreator.pageOrientations.ContainsKey(pageSetup.Orientation))
                {
                    errors.Add("orientationには規定の値を入力して下さい。");
                    foreach (string key in PageOrientationsCreator.pageOrientations.Keys)
                    {
                        errors.Add("[" + key + "] : 設定 - " + PageOrientationsCreator.pageOrientations[key].Description);
                    }
                    break;
                }
            }

            // エラーが有る場合はログ
            if (errors.Count > 0)
            {
                foreach (string error in errors)
                {
                    LogUtil.LogWarn(error);
                }
            }

            return errors.Count == 0;
        }
    }
}
