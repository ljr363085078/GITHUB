using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn_VSTO
{
    public partial class Form1 : Form
    {
        [DllImport("user32.dll", EntryPoint = "FindWindowA")]
        static extern int FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll", EntryPoint = "SetWindowTextA")]
        static extern int SetWindowText(int hwnd, string lpString);
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {

            int h;
            string i = Globals.ThisAddIn.Application.ActiveWorkbook.Name;
            MessageBox.Show(i);
            h = FindWindow("EXCEL7", i);
            MessageBox.Show(h.ToString());
        }
    }
}
