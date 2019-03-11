using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace _02ExcelAddIn
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void toggleButton1_Click(object sender, RibbonControlEventArgs e)
        {
            Share.task1.Visible = this.toggleButton1.Checked;
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Share.task1.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionLeft;
        }

        private void button4_Click(object sender, RibbonControlEventArgs e)
        {
            Share.task1.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionTop;
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            Share.task1.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
        }

        private void button5_Click(object sender, RibbonControlEventArgs e)
        {
            Share.task1.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionBottom;
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            Share.task1.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionFloating;
        }
    }
}
