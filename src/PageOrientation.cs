using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelFormatter.src
{
    class PageOrientation
    {
        public PageOrientation(int value, string description)
        {
            this.Value = value;
            this.Description = description;
        }

        public int Value;
        public string Description;
    }

    class PageOrientationsCreator
    {
        public static Dictionary<string, PageOrientation> pageOrientations = new Dictionary<string, PageOrientation>();

        static PageOrientationsCreator()
        {
            pageOrientations["xlLandscape"] = new PageOrientation(2, "横モード");
            pageOrientations["xlPortrait"] = new PageOrientation(1, "縦モード");
        }
    }
}
