﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace ExcelAddIn_VSTO
{
    public partial class ThisAddIn
    {
        Excel.Application ExcelApp;     //声明一个Excel的应用程序对象
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            ExcelApp = Globals.ThisAddIn.Application;      //获取到加载项所在Excel应用程序
            ExcelApp.ActiveCell.Value = "11";        //测试給活动单元格输入新的值    
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