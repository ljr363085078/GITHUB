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
            checkBox1.Checked = Properties.Settings.Default.if_jgd_open;
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void group1_DialogLauncherClick(object sender, RibbonControlEventArgs e)
        {
            Form1 form1 = new Form1();
            form1.ShowDialog();
        }

        private void checkBox1_Click(object sender, RibbonControlEventArgs e)
        {
            //用复选按钮的状态来控制聚光灯的显示
            if(checkBox1.Checked==true)
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

        private void toggleButton1_Click(object sender, RibbonControlEventArgs e)
        {
            Share.task1.Visible = this.toggleButton1.Checked;
        }

        private void button1_Click_1(object sender, RibbonControlEventArgs e)
        {
            Share.task1.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionLeft;
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            Share.task1.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            Share.task1.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionTop;
        }

        private void button4_Click(object sender, RibbonControlEventArgs e)
        {
            Share.task1.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionBottom;
        }

        private void button5_Click(object sender, RibbonControlEventArgs e)
        {
            Share.task1.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionFloating;
        }
    }
}
