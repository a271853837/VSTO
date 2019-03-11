using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace _02ExcelAddIn
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            UserControl1 uc1 = new UserControl1();
            Share.task1 = Globals.ThisAddIn.CustomTaskPanes.Add(uc1, "工作表导航");

            Share.task1.VisibleChanged += Task1_VisibleChanged;
            Share.task1.DockPositionChanged += Task1_DockPositionChanged;
            Share.task1.Visible = true;

        }

        private void Task1_DockPositionChanged(object sender, EventArgs e)
        {
            Globals.ThisAddIn.Application.StatusBar = Share.task1.DockPosition.ToString();

        }

        private void Task1_VisibleChanged(object sender, EventArgs e)
        {
            Ribbon1 ribbon = Globals.Ribbons.GetRibbon<Ribbon1>();
            ribbon.toggleButton1.Checked = Share.task1.Visible;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
