using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelFormatter.src
{
    class PaperSize
    {
        public PaperSize(int value, string description)
        {
            this.Value = value;
            this.Description = description;
        }

        public int Value { get; set; }
        public string Description { get; set; }
    }

    class PaperSizesCreator
    {
        public static Dictionary<string, PaperSize> paperSizes = new Dictionary<string, PaperSize>();

        static PaperSizesCreator()
        {
            paperSizes["xlPaper10x14"] = new PaperSize(16, "10 x 14 インチ");
            paperSizes["xlPaper11x17"] = new PaperSize(17, "11 x 17 インチ");
            paperSizes["xlPaperA3"] = new PaperSize(8, "A3 (297 mm x 420 mm)");
            paperSizes["xlPaperA4"] = new PaperSize(9, "A4 (210 mm x 297 mm)");
            paperSizes["xlPaperA4Small"] = new PaperSize(10, "A4 スモール (210 mm x 297 mm)");
            paperSizes["xlPaperA5"] = new PaperSize(11, "A5 (148 mm x 210 mm)");
            paperSizes["xlPaperB4"] = new PaperSize(12, "B4 (250 mm x 354 mm)");
            paperSizes["xlPaperB5"] = new PaperSize(13, "A5 (148 mm x 210 mm)");
            paperSizes["xlPaperCsheet"] = new PaperSize(24, "C サイズ シート");
            paperSizes["xlPaperDsheet"] = new PaperSize(25, "D サイズ シート");
            paperSizes["xlPaperEnvelope10"] = new PaperSize(20, "封筒 #10 (4-1/8 x 9-1/2 インチ)");
            paperSizes["xlPaperEnvelope11"] = new PaperSize(21, "封筒 #11 (4-1/2 x 10-3/8 インチ)");
            paperSizes["xlPaperEnvelope12"] = new PaperSize(22, "封筒 #12 (4-1/2 x 11 インチ)");
            paperSizes["xlPaperEnvelope14"] = new PaperSize(23, "封筒 #14 (5 x 11-1/2 インチ)");
            paperSizes["xlPaperEnvelope9"] = new PaperSize(19, "封筒 #9 (3-7/8 x 8-7/2 インチ)");
            paperSizes["xlPaperEnvelopeB4"] = new PaperSize(33, "封筒 B4 (250 mm x 353 mm)");
            paperSizes["xlPaperEnvelopeB5"] = new PaperSize(34, "封筒 B5 (176 mm x 250 mm)");
            paperSizes["xlPaperEnvelopeB6"] = new PaperSize(35, "封筒 B6 (176 mm x 125 mm)");
            paperSizes["xlPaperEnvelopeC3"] = new PaperSize(29, "封筒 C3 (324 mm x 458 mm)");
            paperSizes["xlPaperEnvelopeC4"] = new PaperSize(30, "封筒 C4 (229 mm x 324 mm)");
            paperSizes["xlPaperEnvelopeC5"] = new PaperSize(28, "封筒 C5 (162 mm x 229 mm)");
            paperSizes["xlPaperEnvelopeC6"] = new PaperSize(31, "封筒 C6 (114 mm x 162 mm)");
            paperSizes["xlPaperEnvelopeC65"] = new PaperSize(32, "封筒 C65 (114 mm x 229 mm)");
            paperSizes["xlPaperEnvelopeDL"] = new PaperSize(27, "封筒 DL (110 mm x 220 mm)");
            paperSizes["xlPaperEnvelopeItaly"] = new PaperSize(36, "封筒 (110 mm x 230 mm)");
            paperSizes["xlPaperEnvelopeMonarch"] = new PaperSize(37, "封筒 Monarch (3-7/8 x 7-1/2 インチ)");
            paperSizes["xlPaperEnvelopePersonal"] = new PaperSize(38, "封筒 (3-5/8 x 6-1/2 インチ)");
            paperSizes["xlPaperEsheet"] = new PaperSize(26, "E サイズ シート");
            paperSizes["xlPaperExecutive"] = new PaperSize(7, "エグゼクティブ (7-1/2 x 10-1/2 インチ)");
            paperSizes["xlPaperFanfoldLegalGerman"] = new PaperSize(41, "ドイツ リーガル複写紙 (8-1/2 x 13 インチ)");
            paperSizes["xlPaperFanfoldStdGerman"] = new PaperSize(40, "ドイツ リーガル複写紙 (8-1/2 x 13 インチ)");
            paperSizes["xlPaperFanfoldUS"] = new PaperSize(39, "米国標準複写紙 (14-7/8 x 11 インチ)");
            paperSizes["xlPaperFolio"] = new PaperSize(14, "Folio (8-1/2 x 13 インチ)");
            paperSizes["xlPaperLedger"] = new PaperSize(4, "Ledger (17 x 11 インチ)");
            paperSizes["xlPaperLegal"] = new PaperSize(5, "リーガル (8-1/2 x 14 インチ)");
            paperSizes["xlPaperLetter"] = new PaperSize(1, "レター (8-1/2 x 11 インチ)");
            paperSizes["xlPaperLetterSmall"] = new PaperSize(2, "レター (小) (8-1/2 x 11 インチ)");
            paperSizes["xlPaperNote"] = new PaperSize(18, "ノート (8-1/2 x 11 インチ)");
            paperSizes["xlPaperQuarto"] = new PaperSize(15, "4 つ折版 (215 mm x 275 mm)");
            paperSizes["xlPaperStatement"] = new PaperSize(6, "ステートメント (5-1/2 x 8-1/2 インチ)");
            paperSizes["xlPaperTabloid"] = new PaperSize(3, "タブロイド (11 x 17 インチ)");
            paperSizes["xlPaperUser"] = new PaperSize(256, "ユーザー定義");
        }
    }
}
