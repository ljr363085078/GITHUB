using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace ExcelAddIn_VSTO
{
    
    public partial class ThisAddIn
    {
        Excel.Application ExcelApp;     //声明一个Excel的应用程序对象
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            ExcelApp = Globals.ThisAddIn.Application;      //获取到加载项所在Excel应用程序
            //ExcelApp.ActiveCell.Value = "11";        //测试給活动单元格输入新的值    
            UserControl1 userControl1 = new UserControl1();     //任务栏窗格实例化
            Share.task1 = Globals.ThisAddIn.CustomTaskPanes.Add(userControl1, "工作表导航");
            Share.task1.Visible = true;
            Share.task1.VisibleChanged += new EventHandler(Task1_VisibleChanged);
            Share.task1.DockPositionChanged += Task1_DockPositionChanged;
        }

        private void Task1_DockPositionChanged(object sender, EventArgs e)
        {
            Globals.ThisAddIn.Application.StatusBar = Share.task1.DockPosition.ToString();      //任务窗格改变位置时在excel的状态栏显示位置描述
        }

        private void Task1_VisibleChanged(object sender, EventArgs e)
        {
            Ribbon_VSTO ribbon = Globals.Ribbons.GetRibbon<Ribbon_VSTO>();      //获取功能区可视化设计器中的ribbon
            ribbon.toggleButton1.Checked = Share.task1.Visible;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }
        private void Application_SheetSelectionChange(object Sh, Range Target)
        {
            //throw new NotImplementedException();
            if (Properties.Settings.Default.if_jgd_open == true)
            {
                ExcelApp.Cells.Interior.ColorIndex = 0;
                Range target = Globals.ThisAddIn.Application.Selection;
                if (target.Count > 1)
                {
                    target = target.Cells[1];
                }
                target.EntireColumn.Interior.ColorIndex = 6;
                target.EntireRow.Interior.ColorIndex = 6;
            }
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
            Globals.ThisAddIn.Application.SheetSelectionChange += Application_SheetSelectionChange;
        }
        #endregion
    }
}
