using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace ExcelAddIn_VSTO
{
    public partial class Ribbon_VSTO
    {
        private void Ribbon_VSTO_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            //点击按钮改变if_jgd_open的状态
            if (Properties.Settings.Default.if_jgd_open == false) 
            {
                Properties.Settings.Default.if_jgd_open = true;
                Properties.Settings.Default.Save();
            }
            else
            {
                Properties.Settings.Default.if_jgd_open = false;
                Properties.Settings.Default.Save();
                Globals.ThisAddIn.Application.Cells.Interior.ColorIndex = 0;
            }
        }

        private void group1_DialogLauncherClick(object sender, RibbonControlEventArgs e)
        {
            Form1 form1 = new Form1();
            form1.ShowDialog();
        }
    }
}
