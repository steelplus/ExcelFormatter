using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelFormatter.src
{
    /// <summary>
    /// プログラム実行時オプションをboolで持つクラスです。
    /// </summary>
    class ExecuteOptions
    {
        public bool SelectA1 { get; set; }
        public bool CopyTemplate { get; set; }
        public bool FixPageSetup { get; set; }
        public bool FixFont { get; set; }
        public bool FixHeader { get; set; }
        public bool FixFooter { get; set; }
        public bool RemoveSampleText { get; set; }
        public bool AddSampleText { get; set; }
        public bool OverWriteWriteTime { get; set; }
        public bool SubDirectory { get; set; }
        public string MirrorDirectory { get; set; }

        public bool HasValue()
        {
            return (SelectA1 || CopyTemplate || FixPageSetup || FixFont || FixHeader || FixFooter || RemoveSampleText || AddSampleText || OverWriteWriteTime || SubDirectory || !String.IsNullOrEmpty(MirrorDirectory));
        }
    }
}
