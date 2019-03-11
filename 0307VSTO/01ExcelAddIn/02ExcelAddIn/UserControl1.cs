using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace _02ExcelAddIn
{
    public partial class UserControl1 : UserControl
    {
        public UserControl1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.listBox1.Items.Clear();
            foreach (Microsoft.Office.Interop.Excel.Worksheet item in Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets)
            {
                this.listBox1.Items.Add(item.Name);
            }
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[this.listBox1.Text].Activate();
        }
    }
}
