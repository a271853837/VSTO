using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using _01ExcelAddIn;

namespace _01ExcelAddIn
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Range rng;
            rng = Globals.ThisAddIn.Application.Selection;
            rng.Interior.Color = System.Drawing.Color.Yellow;
        }

        private void group1_DialogLauncherClick(object sender, RibbonControlEventArgs e)
        {
            Form1 fm1 = new Form1();
            fm1.Show();
        }
    }
}
