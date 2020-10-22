namespace ExcelAddIn_VSTO
{
    partial class Ribbon_VSTO : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon_VSTO()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 组件设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            Microsoft.Office.Tools.Ribbon.RibbonDialogLauncher ribbonDialogLauncherImpl1 = this.Factory.CreateRibbonDialogLauncher();
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.button_jgd = this.Factory.CreateRibbonButton();
            this.checkBox1 = this.Factory.CreateRibbonCheckBox();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.toggleButton1 = this.Factory.CreateRibbonToggleButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.button_left = this.Factory.CreateRibbonButton();
            this.button_right = this.Factory.CreateRibbonButton();
            this.button_up = this.Factory.CreateRibbonButton();
            this.button_down = this.Factory.CreateRibbonButton();
            this.button_floating = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Label = "tab1";
            this.tab1.Name = "tab1";
            this.tab1.Position = this.Factory.RibbonPosition.BeforeOfficeId("TabInsert");
            // 
            // group1
            // 
            this.group1.DialogLauncher = ribbonDialogLauncherImpl1;
            this.group1.Items.Add(this.button_jgd);
            this.group1.Items.Add(this.checkBox1);
            this.group1.Label = "聚光灯";
            this.group1.Name = "group1";
            this.group1.DialogLauncherClick += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.group1_DialogLauncherClick);
            // 
            // button_jgd
            // 
            this.button_jgd.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_jgd.Label = "聚光灯";
            this.button_jgd.Name = "button_jgd";
            this.button_jgd.OfficeImageId = "Call";
            this.button_jgd.ShowImage = true;
            this.button_jgd.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // checkBox1
            // 
            this.checkBox1.Label = "聚光灯";
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.checkBox1_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.toggleButton1);
            this.group2.Label = "可见性";
            this.group2.Name = "group2";
            // 
            // toggleButton1
            // 
            this.toggleButton1.Label = "显示|隐藏";
            this.toggleButton1.Name = "toggleButton1";
            this.toggleButton1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.toggleButton1_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.button_left);
            this.group3.Items.Add(this.button_right);
            this.group3.Items.Add(this.button_up);
            this.group3.Items.Add(this.button_down);
            this.group3.Items.Add(this.button_floating);
            this.group3.Label = "任务窗格停靠位置";
            this.group3.Name = "group3";
            // 
            // button_left
            // 
            this.button_left.Label = "左";
            this.button_left.Name = "button_left";
            this.button_left.OfficeImageId = "L";
            this.button_left.ShowImage = true;
            this.button_left.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click_1);
            // 
            // button_right
            // 
            this.button_right.Label = "右";
            this.button_right.Name = "button_right";
            this.button_right.OfficeImageId = "R";
            this.button_right.ShowImage = true;
            this.button_right.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button2_Click);
            // 
            // button_up
            // 
            this.button_up.Label = "上";
            this.button_up.Name = "button_up";
            this.button_up.OfficeImageId = "U";
            this.button_up.ShowImage = true;
            this.button_up.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button3_Click);
            // 
            // button_down
            // 
            this.button_down.Label = "下";
            this.button_down.Name = "button_down";
            this.button_down.OfficeImageId = "D";
            this.button_down.ShowImage = true;
            this.button_down.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button4_Click);
            // 
            // button_floating
            // 
            this.button_floating.Label = "浮动";
            this.button_floating.Name = "button_floating";
            this.button_floating.OfficeImageId = "F";
            this.button_floating.ShowImage = true;
            this.button_floating.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button5_Click);
            // 
            // Ribbon_VSTO
            // 
            this.Name = "Ribbon_VSTO";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_VSTO_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_jgd;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBox1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleButton1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_left;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_right;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_up;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_down;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_floating;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon_VSTO Ribbon_VSTO
        {
            get { return this.GetRibbon<Ribbon_VSTO>(); }
        }
    }
}
