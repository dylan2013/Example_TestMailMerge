using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA;
using FISCA.Presentation;

namespace TestMailMerge
{
    public class Program
    {
        [MainMethod()]
        static public void Main()
        {

            RibbonBarItem totle = FISCA.Presentation.MotherForm.RibbonBarItems["學生", "其它"];
            totle["列印學校基本資料"].Click += delegate
            {
                Form f = new Form();
                f.ShowDialog();
            };

        }
    }
}
