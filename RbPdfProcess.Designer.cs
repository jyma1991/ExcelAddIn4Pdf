namespace ExcelAddIn4Pdf
{
    partial class RbPdfProcess : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RbPdfProcess()
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
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(RbPdfProcess));
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl2 = this.Factory.CreateRibbonDropDownItem();
            this.tab1 = this.Factory.CreateRibbonTab();
            this.gpPanel = this.Factory.CreateRibbonGroup();
            this.btnOpen = this.Factory.CreateRibbonButton();
            this.dropDown = this.Factory.CreateRibbonDropDown();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnCreate = this.Factory.CreateRibbonButton();
            this.btnSave = this.Factory.CreateRibbonButton();
            this.btnInsertCount = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.gpPanel.SuspendLayout();
            this.group1.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.gpPanel);
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // gpPanel
            // 
            this.gpPanel.Items.Add(this.btnOpen);
            this.gpPanel.Items.Add(this.dropDown);
            this.gpPanel.Items.Add(this.btnInsertCount);
            this.gpPanel.Label = "设置";
            this.gpPanel.Name = "gpPanel";
            // 
            // btnOpen
            // 
            this.btnOpen.Image = ((System.Drawing.Image)(resources.GetObject("btnOpen.Image")));
            this.btnOpen.Label = "行程单选择";
            this.btnOpen.Name = "btnOpen";
            this.btnOpen.ShowImage = true;
            this.btnOpen.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnOpen_Click);
            // 
            // dropDown
            // 
            ribbonDropDownItemImpl1.Label = "Charter";
            ribbonDropDownItemImpl2.Label = "Driver Guide";
            this.dropDown.Items.Add(ribbonDropDownItemImpl1);
            this.dropDown.Items.Add(ribbonDropDownItemImpl2);
            this.dropDown.Label = "Charter";
            this.dropDown.Name = "dropDown";
            this.dropDown.ShowLabel = false;
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnCreate);
            this.group1.Items.Add(this.btnSave);
            this.group1.Label = "报价单生成";
            this.group1.Name = "group1";
            // 
            // btnCreate
            // 
            this.btnCreate.Image = ((System.Drawing.Image)(resources.GetObject("btnCreate.Image")));
            this.btnCreate.Label = "预览报价单";
            this.btnCreate.Name = "btnCreate";
            this.btnCreate.ShowImage = true;
            this.btnCreate.Visible = false;
            this.btnCreate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCreate_Click);
            // 
            // btnSave
            // 
            this.btnSave.Image = ((System.Drawing.Image)(resources.GetObject("btnSave.Image")));
            this.btnSave.Label = "保存报价单";
            this.btnSave.Name = "btnSave";
            this.btnSave.ShowImage = true;
            this.btnSave.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSave_Click);
            // 
            // btnInsertCount
            // 
            this.btnInsertCount.Image = ((System.Drawing.Image)(resources.GetObject("btnInsertCount.Image")));
            this.btnInsertCount.Label = "插入折扣";
            this.btnInsertCount.Name = "btnInsertCount";
            this.btnInsertCount.ShowImage = true;
            this.btnInsertCount.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnInsertCount_Click);
            // 
            // RbPdfProcess
            // 
            this.Name = "RbPdfProcess";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RbPdfProcess_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.gpPanel.ResumeLayout(false);
            this.gpPanel.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup gpPanel;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnOpen;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCreate;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown dropDown;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSave;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnInsertCount;
    }

    partial class ThisRibbonCollection
    {
        internal RbPdfProcess RbPdfProcess
        {
            get { return this.GetRibbon<RbPdfProcess>(); }
        }
    }
}
