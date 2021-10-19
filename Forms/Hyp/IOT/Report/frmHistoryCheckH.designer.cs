namespace BMS
{
    partial class frmHistoryCheckH
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
			this.components = new System.ComponentModel.Container();
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmHistoryCheckH));
			this.grdData = new DevExpress.XtraGrid.GridControl();
			this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
			this.btnShowHistoryCheckOrderDetailH = new System.Windows.Forms.ToolStripMenuItem();
			this.nGHELẠIFILEÂMTHANHToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.nGHEÂMTHANHHYPONICSẢNXUẤTLẦN1ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.nGHEÂMTHANHCHIỀUFToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.nGHEÂMTHANHCHIỀURToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.nGHEÂMTHANHHYPONICSỬAHÀNGKHIHÀNGNGToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.nGHEÂMTHANHCHIỀUFToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
			this.nGHEÂMTHANHCHIỀURToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
			this.nGHEÂMTHANHHYPONICKIỂMTRAHÀNGKHIKHÁCHHÀNGTRẢVỀToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.nGHEÂMTHANHCHIỀUFToolStripMenuItem2 = new System.Windows.Forms.ToolStripMenuItem();
			this.nGHEÂMTHANHCHIỀURToolStripMenuItem2 = new System.Windows.Forms.ToolStripMenuItem();
			this.nGHEÂMTHANHHYPONICSAUKHISỬALỖIKHÁCHHÀNGTRẢVỀToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.nGHEÂMTHANHCHIỀUFToolStripMenuItem3 = new System.Windows.Forms.ToolStripMenuItem();
			this.nGHEÂMTHANHCHIỀURToolStripMenuItem3 = new System.Windows.Forms.ToolStripMenuItem();
			this.nGHEÂMTHANHToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.nGHEÂMTHANHHYPONICTHEOQRCODERToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.xEMẢNHHYPONICTHEOQRCODEToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.grvData = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.colPID = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colSTT = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colGroup = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colID = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colQRCode = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colOrderCode = new DevExpress.XtraGrid.Columns.GridColumn();
			this.repositoryItemMemoEdit1 = new DevExpress.XtraEditors.Repository.RepositoryItemMemoEdit();
			this.colMonitor = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colDateLR = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colMain = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colSalesOrder = new DevExpress.XtraGrid.Columns.GridColumn();
			this.repositoryItemCheckEdit1 = new DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit();
			this.repositoryItemCheckEdit2 = new DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit();
			this.label1 = new System.Windows.Forms.Label();
			this.dtpFromDate = new System.Windows.Forms.DateTimePicker();
			this.label2 = new System.Windows.Forms.Label();
			this.dtpEndDate = new System.Windows.Forms.DateTimePicker();
			this.label3 = new System.Windows.Forms.Label();
			this.txtFindText = new System.Windows.Forms.TextBox();
			this.btnFind = new System.Windows.Forms.Button();
			this.btnExportExecl = new DevExpress.XtraEditors.SimpleButton();
			this.sfdSaveDialog = new System.Windows.Forms.SaveFileDialog();
			((System.ComponentModel.ISupportInitialize)(this.grdData)).BeginInit();
			this.contextMenuStrip1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.grvData)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemMemoEdit1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemCheckEdit1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemCheckEdit2)).BeginInit();
			this.SuspendLayout();
			// 
			// grdData
			// 
			this.grdData.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.grdData.ContextMenuStrip = this.contextMenuStrip1;
			this.grdData.Font = new System.Drawing.Font("Times New Roman", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.grdData.Location = new System.Drawing.Point(3, 49);
			this.grdData.MainView = this.grvData;
			this.grdData.Name = "grdData";
			this.grdData.RepositoryItems.AddRange(new DevExpress.XtraEditors.Repository.RepositoryItem[] {
            this.repositoryItemCheckEdit1,
            this.repositoryItemMemoEdit1,
            this.repositoryItemCheckEdit2});
			this.grdData.Size = new System.Drawing.Size(1018, 485);
			this.grdData.TabIndex = 4;
			this.grdData.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.grvData});
			// 
			// contextMenuStrip1
			// 
			this.contextMenuStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
			this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.btnShowHistoryCheckOrderDetailH,
            this.nGHELẠIFILEÂMTHANHToolStripMenuItem,
            this.nGHEÂMTHANHToolStripMenuItem,
            this.nGHEÂMTHANHHYPONICTHEOQRCODERToolStripMenuItem,
            this.xEMẢNHHYPONICTHEOQRCODEToolStripMenuItem});
			this.contextMenuStrip1.Name = "contextMenuStrip1";
			this.contextMenuStrip1.Size = new System.Drawing.Size(321, 114);
			// 
			// btnShowHistoryCheckOrderDetailH
			// 
			this.btnShowHistoryCheckOrderDetailH.Name = "btnShowHistoryCheckOrderDetailH";
			this.btnShowHistoryCheckOrderDetailH.Size = new System.Drawing.Size(320, 22);
			this.btnShowHistoryCheckOrderDetailH.Tag = "Module_Film";
			this.btnShowHistoryCheckOrderDetailH.Text = "CHI TIẾT KIẾM TRA THEO ORDER";
			this.btnShowHistoryCheckOrderDetailH.Click += new System.EventHandler(this.btnShowWorking_Click);
			// 
			// nGHELẠIFILEÂMTHANHToolStripMenuItem
			// 
			this.nGHELẠIFILEÂMTHANHToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.nGHEÂMTHANHHYPONICSẢNXUẤTLẦN1ToolStripMenuItem,
            this.nGHEÂMTHANHHYPONICSỬAHÀNGKHIHÀNGNGToolStripMenuItem,
            this.nGHEÂMTHANHHYPONICKIỂMTRAHÀNGKHIKHÁCHHÀNGTRẢVỀToolStripMenuItem,
            this.nGHEÂMTHANHHYPONICSAUKHISỬALỖIKHÁCHHÀNGTRẢVỀToolStripMenuItem});
			this.nGHELẠIFILEÂMTHANHToolStripMenuItem.Name = "nGHELẠIFILEÂMTHANHToolStripMenuItem";
			this.nGHELẠIFILEÂMTHANHToolStripMenuItem.Size = new System.Drawing.Size(320, 22);
			this.nGHELẠIFILEÂMTHANHToolStripMenuItem.Text = "NGHE LẠI FILE ÂM THANH ";
			// 
			// nGHEÂMTHANHHYPONICSẢNXUẤTLẦN1ToolStripMenuItem
			// 
			this.nGHEÂMTHANHHYPONICSẢNXUẤTLẦN1ToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.nGHEÂMTHANHCHIỀUFToolStripMenuItem,
            this.nGHEÂMTHANHCHIỀURToolStripMenuItem});
			this.nGHEÂMTHANHHYPONICSẢNXUẤTLẦN1ToolStripMenuItem.Name = "nGHEÂMTHANHHYPONICSẢNXUẤTLẦN1ToolStripMenuItem";
			this.nGHEÂMTHANHHYPONICSẢNXUẤTLẦN1ToolStripMenuItem.Size = new System.Drawing.Size(461, 22);
			this.nGHEÂMTHANHHYPONICSẢNXUẤTLẦN1ToolStripMenuItem.Text = "NGHE ÂM THANH HYPONIC SẢN XUẤT LẦN 1";
			// 
			// nGHEÂMTHANHCHIỀUFToolStripMenuItem
			// 
			this.nGHEÂMTHANHCHIỀUFToolStripMenuItem.Name = "nGHEÂMTHANHCHIỀUFToolStripMenuItem";
			this.nGHEÂMTHANHCHIỀUFToolStripMenuItem.Size = new System.Drawing.Size(219, 22);
			this.nGHEÂMTHANHCHIỀUFToolStripMenuItem.Text = "NGHE ÂM THANH CHIỀU F";
			this.nGHEÂMTHANHCHIỀUFToolStripMenuItem.Click += new System.EventHandler(this.nGHEÂMTHANHCHIỀUFToolStripMenuItem_Click);
			// 
			// nGHEÂMTHANHCHIỀURToolStripMenuItem
			// 
			this.nGHEÂMTHANHCHIỀURToolStripMenuItem.Name = "nGHEÂMTHANHCHIỀURToolStripMenuItem";
			this.nGHEÂMTHANHCHIỀURToolStripMenuItem.Size = new System.Drawing.Size(219, 22);
			this.nGHEÂMTHANHCHIỀURToolStripMenuItem.Text = "NGHE ÂM THANH CHIỀU R";
			this.nGHEÂMTHANHCHIỀURToolStripMenuItem.Click += new System.EventHandler(this.nGHEÂMTHANHCHIỀURToolStripMenuItem_Click);
			// 
			// nGHEÂMTHANHHYPONICSỬAHÀNGKHIHÀNGNGToolStripMenuItem
			// 
			this.nGHEÂMTHANHHYPONICSỬAHÀNGKHIHÀNGNGToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.nGHEÂMTHANHCHIỀUFToolStripMenuItem1,
            this.nGHEÂMTHANHCHIỀURToolStripMenuItem1});
			this.nGHEÂMTHANHHYPONICSỬAHÀNGKHIHÀNGNGToolStripMenuItem.Name = "nGHEÂMTHANHHYPONICSỬAHÀNGKHIHÀNGNGToolStripMenuItem";
			this.nGHEÂMTHANHHYPONICSỬAHÀNGKHIHÀNGNGToolStripMenuItem.Size = new System.Drawing.Size(461, 22);
			this.nGHEÂMTHANHHYPONICSỬAHÀNGKHIHÀNGNGToolStripMenuItem.Text = "NGHE ÂM THANH HYPONIC SỬA HÀNG KHI HÀNG NG";
			// 
			// nGHEÂMTHANHCHIỀUFToolStripMenuItem1
			// 
			this.nGHEÂMTHANHCHIỀUFToolStripMenuItem1.Name = "nGHEÂMTHANHCHIỀUFToolStripMenuItem1";
			this.nGHEÂMTHANHCHIỀUFToolStripMenuItem1.Size = new System.Drawing.Size(219, 22);
			this.nGHEÂMTHANHCHIỀUFToolStripMenuItem1.Text = "NGHE ÂM THANH CHIỀU F";
			this.nGHEÂMTHANHCHIỀUFToolStripMenuItem1.Click += new System.EventHandler(this.nGHEÂMTHANHCHIỀUFToolStripMenuItem1_Click);
			// 
			// nGHEÂMTHANHCHIỀURToolStripMenuItem1
			// 
			this.nGHEÂMTHANHCHIỀURToolStripMenuItem1.Name = "nGHEÂMTHANHCHIỀURToolStripMenuItem1";
			this.nGHEÂMTHANHCHIỀURToolStripMenuItem1.Size = new System.Drawing.Size(219, 22);
			this.nGHEÂMTHANHCHIỀURToolStripMenuItem1.Text = "NGHE ÂM THANH CHIỀU R";
			this.nGHEÂMTHANHCHIỀURToolStripMenuItem1.Click += new System.EventHandler(this.nGHEÂMTHANHCHIỀURToolStripMenuItem1_Click);
			// 
			// nGHEÂMTHANHHYPONICKIỂMTRAHÀNGKHIKHÁCHHÀNGTRẢVỀToolStripMenuItem
			// 
			this.nGHEÂMTHANHHYPONICKIỂMTRAHÀNGKHIKHÁCHHÀNGTRẢVỀToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.nGHEÂMTHANHCHIỀUFToolStripMenuItem2,
            this.nGHEÂMTHANHCHIỀURToolStripMenuItem2});
			this.nGHEÂMTHANHHYPONICKIỂMTRAHÀNGKHIKHÁCHHÀNGTRẢVỀToolStripMenuItem.Name = "nGHEÂMTHANHHYPONICKIỂMTRAHÀNGKHIKHÁCHHÀNGTRẢVỀToolStripMenuItem";
			this.nGHEÂMTHANHHYPONICKIỂMTRAHÀNGKHIKHÁCHHÀNGTRẢVỀToolStripMenuItem.Size = new System.Drawing.Size(461, 22);
			this.nGHEÂMTHANHHYPONICKIỂMTRAHÀNGKHIKHÁCHHÀNGTRẢVỀToolStripMenuItem.Text = "NGHE ÂM THANH HYPONIC KIỂM TRA HÀNG KHI KHÁCH HÀNG TRẢ VỀ";
			// 
			// nGHEÂMTHANHCHIỀUFToolStripMenuItem2
			// 
			this.nGHEÂMTHANHCHIỀUFToolStripMenuItem2.Name = "nGHEÂMTHANHCHIỀUFToolStripMenuItem2";
			this.nGHEÂMTHANHCHIỀUFToolStripMenuItem2.Size = new System.Drawing.Size(219, 22);
			this.nGHEÂMTHANHCHIỀUFToolStripMenuItem2.Text = "NGHE ÂM THANH CHIỀU F";
			this.nGHEÂMTHANHCHIỀUFToolStripMenuItem2.Click += new System.EventHandler(this.nGHEÂMTHANHCHIỀUFToolStripMenuItem2_Click);
			// 
			// nGHEÂMTHANHCHIỀURToolStripMenuItem2
			// 
			this.nGHEÂMTHANHCHIỀURToolStripMenuItem2.Name = "nGHEÂMTHANHCHIỀURToolStripMenuItem2";
			this.nGHEÂMTHANHCHIỀURToolStripMenuItem2.Size = new System.Drawing.Size(219, 22);
			this.nGHEÂMTHANHCHIỀURToolStripMenuItem2.Text = "NGHE ÂM THANH CHIỀU R";
			this.nGHEÂMTHANHCHIỀURToolStripMenuItem2.Click += new System.EventHandler(this.nGHEÂMTHANHCHIỀURToolStripMenuItem2_Click);
			// 
			// nGHEÂMTHANHHYPONICSAUKHISỬALỖIKHÁCHHÀNGTRẢVỀToolStripMenuItem
			// 
			this.nGHEÂMTHANHHYPONICSAUKHISỬALỖIKHÁCHHÀNGTRẢVỀToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.nGHEÂMTHANHCHIỀUFToolStripMenuItem3,
            this.nGHEÂMTHANHCHIỀURToolStripMenuItem3});
			this.nGHEÂMTHANHHYPONICSAUKHISỬALỖIKHÁCHHÀNGTRẢVỀToolStripMenuItem.Name = "nGHEÂMTHANHHYPONICSAUKHISỬALỖIKHÁCHHÀNGTRẢVỀToolStripMenuItem";
			this.nGHEÂMTHANHHYPONICSAUKHISỬALỖIKHÁCHHÀNGTRẢVỀToolStripMenuItem.Size = new System.Drawing.Size(461, 22);
			this.nGHEÂMTHANHHYPONICSAUKHISỬALỖIKHÁCHHÀNGTRẢVỀToolStripMenuItem.Text = "NGHE ÂM THANH HYPONIC SAU KHI SỬA LỖI KHÁCH HÀNG TRẢ VỀ";
			// 
			// nGHEÂMTHANHCHIỀUFToolStripMenuItem3
			// 
			this.nGHEÂMTHANHCHIỀUFToolStripMenuItem3.Name = "nGHEÂMTHANHCHIỀUFToolStripMenuItem3";
			this.nGHEÂMTHANHCHIỀUFToolStripMenuItem3.Size = new System.Drawing.Size(219, 22);
			this.nGHEÂMTHANHCHIỀUFToolStripMenuItem3.Text = "NGHE ÂM THANH CHIỀU F";
			this.nGHEÂMTHANHCHIỀUFToolStripMenuItem3.Click += new System.EventHandler(this.nGHEÂMTHANHCHIỀUFToolStripMenuItem3_Click);
			// 
			// nGHEÂMTHANHCHIỀURToolStripMenuItem3
			// 
			this.nGHEÂMTHANHCHIỀURToolStripMenuItem3.Name = "nGHEÂMTHANHCHIỀURToolStripMenuItem3";
			this.nGHEÂMTHANHCHIỀURToolStripMenuItem3.Size = new System.Drawing.Size(219, 22);
			this.nGHEÂMTHANHCHIỀURToolStripMenuItem3.Text = "NGHE ÂM THANH CHIỀU R";
			this.nGHEÂMTHANHCHIỀURToolStripMenuItem3.Click += new System.EventHandler(this.nGHEÂMTHANHCHIỀURToolStripMenuItem3_Click);
			// 
			// nGHEÂMTHANHToolStripMenuItem
			// 
			this.nGHEÂMTHANHToolStripMenuItem.Name = "nGHEÂMTHANHToolStripMenuItem";
			this.nGHEÂMTHANHToolStripMenuItem.Size = new System.Drawing.Size(320, 22);
			this.nGHEÂMTHANHToolStripMenuItem.Text = "NGHE ÂM THANH HYPONIC THEO QRCODE F";
			this.nGHEÂMTHANHToolStripMenuItem.Click += new System.EventHandler(this.nGHEÂMTHANHToolStripMenuItem_Click);
			// 
			// nGHEÂMTHANHHYPONICTHEOQRCODERToolStripMenuItem
			// 
			this.nGHEÂMTHANHHYPONICTHEOQRCODERToolStripMenuItem.Name = "nGHEÂMTHANHHYPONICTHEOQRCODERToolStripMenuItem";
			this.nGHEÂMTHANHHYPONICTHEOQRCODERToolStripMenuItem.Size = new System.Drawing.Size(320, 22);
			this.nGHEÂMTHANHHYPONICTHEOQRCODERToolStripMenuItem.Text = "NGHE ÂM THANH HYPONIC THEO QRCODE R";
			this.nGHEÂMTHANHHYPONICTHEOQRCODERToolStripMenuItem.Click += new System.EventHandler(this.nGHEÂMTHANHHYPONICTHEOQRCODERToolStripMenuItem_Click);
			// 
			// xEMẢNHHYPONICTHEOQRCODEToolStripMenuItem
			// 
			this.xEMẢNHHYPONICTHEOQRCODEToolStripMenuItem.Name = "xEMẢNHHYPONICTHEOQRCODEToolStripMenuItem";
			this.xEMẢNHHYPONICTHEOQRCODEToolStripMenuItem.Size = new System.Drawing.Size(320, 22);
			this.xEMẢNHHYPONICTHEOQRCODEToolStripMenuItem.Text = "ẢNH SẢN PHẨM BƠM MỠ";
			this.xEMẢNHHYPONICTHEOQRCODEToolStripMenuItem.Click += new System.EventHandler(this.xEMẢNHHYPONICTHEOQRCODEToolStripMenuItem_Click);
			// 
			// grvData
			// 
			this.grvData.Appearance.FocusedRow.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(192)))), ((int)(((byte)(255)))));
			this.grvData.Appearance.FocusedRow.BackColor2 = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(192)))), ((int)(((byte)(255)))));
			this.grvData.Appearance.FocusedRow.Options.UseBackColor = true;
			this.grvData.Appearance.HeaderPanel.Options.UseTextOptions = true;
			this.grvData.Appearance.HeaderPanel.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.grvData.Appearance.HideSelectionRow.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(192)))), ((int)(((byte)(255)))));
			this.grvData.Appearance.HideSelectionRow.BackColor2 = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(192)))), ((int)(((byte)(255)))));
			this.grvData.Appearance.HideSelectionRow.Options.UseBackColor = true;
			this.grvData.Appearance.SelectedRow.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(192)))), ((int)(((byte)(255)))));
			this.grvData.Appearance.SelectedRow.BackColor2 = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(192)))), ((int)(((byte)(255)))));
			this.grvData.Appearance.SelectedRow.Options.UseBackColor = true;
			this.grvData.ColumnPanelRowHeight = 40;
			this.grvData.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
            this.colPID,
            this.colSTT,
            this.colGroup,
            this.colID,
            this.colQRCode,
            this.colOrderCode,
            this.colMonitor,
            this.colDateLR,
            this.colMain,
            this.colSalesOrder});
			this.grvData.GridControl = this.grdData;
			this.grvData.HorzScrollStep = 5;
			this.grvData.Name = "grvData";
			this.grvData.OptionsBehavior.AutoExpandAllGroups = true;
			this.grvData.OptionsCustomization.AllowRowSizing = true;
			this.grvData.OptionsFind.AlwaysVisible = true;
			this.grvData.OptionsFind.ShowCloseButton = false;
			this.grvData.OptionsSelection.MultiSelect = true;
			this.grvData.OptionsView.ColumnAutoWidth = false;
			this.grvData.OptionsView.RowAutoHeight = true;
			this.grvData.OptionsView.ShowFooter = true;
			this.grvData.OptionsView.ShowGroupPanel = false;
			this.grvData.DoubleClick += new System.EventHandler(this.grvData_DoubleClick);
			// 
			// colPID
			// 
			this.colPID.AppearanceCell.Options.UseTextOptions = true;
			this.colPID.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colPID.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colPID.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colPID.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colPID.AppearanceHeader.Options.UseBackColor = true;
			this.colPID.AppearanceHeader.Options.UseFont = true;
			this.colPID.AppearanceHeader.Options.UseTextOptions = true;
			this.colPID.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colPID.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colPID.Caption = "PID";
			this.colPID.FieldName = "ProductCode";
			this.colPID.Name = "colPID";
			this.colPID.OptionsColumn.AllowEdit = false;
			this.colPID.OptionsColumn.AllowSize = false;
			this.colPID.Visible = true;
			this.colPID.VisibleIndex = 4;
			this.colPID.Width = 210;
			// 
			// colSTT
			// 
			this.colSTT.AppearanceCell.Options.UseTextOptions = true;
			this.colSTT.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colSTT.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colSTT.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colSTT.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colSTT.AppearanceHeader.Options.UseBackColor = true;
			this.colSTT.AppearanceHeader.Options.UseFont = true;
			this.colSTT.AppearanceHeader.Options.UseTextOptions = true;
			this.colSTT.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colSTT.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colSTT.Caption = "STT";
			this.colSTT.FieldName = "STT";
			this.colSTT.Name = "colSTT";
			this.colSTT.OptionsColumn.AllowEdit = false;
			this.colSTT.Summary.AddRange(new DevExpress.XtraGrid.GridSummaryItem[] {
            new DevExpress.XtraGrid.GridColumnSummaryItem(DevExpress.Data.SummaryItemType.Count)});
			this.colSTT.Visible = true;
			this.colSTT.VisibleIndex = 0;
			this.colSTT.Width = 101;
			// 
			// colGroup
			// 
			this.colGroup.AppearanceCell.Options.UseTextOptions = true;
			this.colGroup.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colGroup.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colGroup.AppearanceHeader.Options.UseBackColor = true;
			this.colGroup.AppearanceHeader.Options.UseTextOptions = true;
			this.colGroup.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colGroup.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colGroup.Caption = "Order Code";
			this.colGroup.FieldName = "OrderCode";
			this.colGroup.Name = "colGroup";
			this.colGroup.OptionsColumn.AllowEdit = false;
			// 
			// colID
			// 
			this.colID.AppearanceCell.Options.UseTextOptions = true;
			this.colID.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colID.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colID.AppearanceHeader.Options.UseBackColor = true;
			this.colID.AppearanceHeader.Options.UseTextOptions = true;
			this.colID.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colID.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colID.Caption = "ID";
			this.colID.FieldName = "ID";
			this.colID.Name = "colID";
			this.colID.OptionsColumn.AllowEdit = false;
			// 
			// colQRCode
			// 
			this.colQRCode.AppearanceCell.Options.UseTextOptions = true;
			this.colQRCode.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colQRCode.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colQRCode.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colQRCode.AppearanceHeader.Options.UseBackColor = true;
			this.colQRCode.AppearanceHeader.Options.UseFont = true;
			this.colQRCode.AppearanceHeader.Options.UseTextOptions = true;
			this.colQRCode.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colQRCode.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colQRCode.Caption = "QRCode";
			this.colQRCode.FieldName = "QRCode";
			this.colQRCode.Name = "colQRCode";
			this.colQRCode.OptionsColumn.ReadOnly = true;
			this.colQRCode.Visible = true;
			this.colQRCode.VisibleIndex = 2;
			this.colQRCode.Width = 219;
			// 
			// colOrderCode
			// 
			this.colOrderCode.AppearanceCell.Options.UseTextOptions = true;
			this.colOrderCode.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colOrderCode.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colOrderCode.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colOrderCode.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colOrderCode.AppearanceHeader.Options.UseBackColor = true;
			this.colOrderCode.AppearanceHeader.Options.UseFont = true;
			this.colOrderCode.AppearanceHeader.Options.UseTextOptions = true;
			this.colOrderCode.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colOrderCode.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colOrderCode.Caption = "Order Code";
			this.colOrderCode.ColumnEdit = this.repositoryItemMemoEdit1;
			this.colOrderCode.FieldName = "OrderCode";
			this.colOrderCode.Name = "colOrderCode";
			this.colOrderCode.OptionsColumn.AllowSize = false;
			this.colOrderCode.OptionsColumn.ReadOnly = true;
			this.colOrderCode.Visible = true;
			this.colOrderCode.VisibleIndex = 3;
			this.colOrderCode.Width = 238;
			// 
			// repositoryItemMemoEdit1
			// 
			this.repositoryItemMemoEdit1.Appearance.Options.UseTextOptions = true;
			this.repositoryItemMemoEdit1.Appearance.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.repositoryItemMemoEdit1.Name = "repositoryItemMemoEdit1";
			// 
			// colMonitor
			// 
			this.colMonitor.AppearanceCell.Options.UseTextOptions = true;
			this.colMonitor.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colMonitor.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colMonitor.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colMonitor.AppearanceHeader.Options.UseBackColor = true;
			this.colMonitor.AppearanceHeader.Options.UseFont = true;
			this.colMonitor.AppearanceHeader.Options.UseTextOptions = true;
			this.colMonitor.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colMonitor.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colMonitor.Caption = "Người phụ trách";
			this.colMonitor.FieldName = "Monitor";
			this.colMonitor.Name = "colMonitor";
			this.colMonitor.OptionsColumn.AllowEdit = false;
			this.colMonitor.OptionsColumn.AllowSize = false;
			this.colMonitor.Width = 168;
			// 
			// colDateLR
			// 
			this.colDateLR.AppearanceCell.Options.UseTextOptions = true;
			this.colDateLR.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colDateLR.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colDateLR.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colDateLR.AppearanceHeader.Options.UseBackColor = true;
			this.colDateLR.AppearanceHeader.Options.UseFont = true;
			this.colDateLR.AppearanceHeader.Options.UseTextOptions = true;
			this.colDateLR.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colDateLR.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colDateLR.Caption = "Thời gian";
			this.colDateLR.DisplayFormat.FormatString = "dd/MM/yyyy HH:mm:ss";
			this.colDateLR.DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
			this.colDateLR.FieldName = "DateLR";
			this.colDateLR.Name = "colDateLR";
			this.colDateLR.OptionsColumn.AllowEdit = false;
			this.colDateLR.Visible = true;
			this.colDateLR.VisibleIndex = 1;
			this.colDateLR.Width = 186;
			// 
			// colMain
			// 
			this.colMain.AppearanceCell.Options.UseTextOptions = true;
			this.colMain.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colMain.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colMain.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colMain.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colMain.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colMain.AppearanceHeader.Options.UseBackColor = true;
			this.colMain.AppearanceHeader.Options.UseFont = true;
			this.colMain.AppearanceHeader.Options.UseTextOptions = true;
			this.colMain.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colMain.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colMain.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colMain.Caption = "Line";
			this.colMain.FieldName = "Mains";
			this.colMain.Name = "colMain";
			this.colMain.OptionsColumn.AllowEdit = false;
			this.colMain.OptionsColumn.ReadOnly = true;
			this.colMain.UnboundType = DevExpress.Data.UnboundColumnType.String;
			// 
			// colSalesOrder
			// 
			this.colSalesOrder.AppearanceCell.Options.UseFont = true;
			this.colSalesOrder.AppearanceCell.Options.UseTextOptions = true;
			this.colSalesOrder.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colSalesOrder.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colSalesOrder.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colSalesOrder.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colSalesOrder.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colSalesOrder.AppearanceHeader.Options.UseBackColor = true;
			this.colSalesOrder.AppearanceHeader.Options.UseFont = true;
			this.colSalesOrder.AppearanceHeader.Options.UseTextOptions = true;
			this.colSalesOrder.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colSalesOrder.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colSalesOrder.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colSalesOrder.Caption = "Sales Order";
			this.colSalesOrder.FieldName = "SalesOrder";
			this.colSalesOrder.Name = "colSalesOrder";
			this.colSalesOrder.OptionsColumn.ReadOnly = true;
			this.colSalesOrder.Visible = true;
			this.colSalesOrder.VisibleIndex = 5;
			this.colSalesOrder.Width = 182;
			// 
			// repositoryItemCheckEdit1
			// 
			this.repositoryItemCheckEdit1.AutoHeight = false;
			this.repositoryItemCheckEdit1.Caption = "Check";
			this.repositoryItemCheckEdit1.DisplayValueChecked = "\"1\"";
			this.repositoryItemCheckEdit1.DisplayValueUnchecked = "\"0\"";
			this.repositoryItemCheckEdit1.EditValueChangedFiringMode = DevExpress.XtraEditors.Controls.EditValueChangedFiringMode.Buffered;
			this.repositoryItemCheckEdit1.Name = "repositoryItemCheckEdit1";
			this.repositoryItemCheckEdit1.NullStyle = DevExpress.XtraEditors.Controls.StyleIndeterminate.Unchecked;
			// 
			// repositoryItemCheckEdit2
			// 
			this.repositoryItemCheckEdit2.AutoHeight = false;
			this.repositoryItemCheckEdit2.Caption = "Check";
			this.repositoryItemCheckEdit2.Name = "repositoryItemCheckEdit2";
			this.repositoryItemCheckEdit2.NullStyle = DevExpress.XtraEditors.Controls.StyleIndeterminate.Unchecked;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(21, 18);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(46, 13);
			this.label1.TabIndex = 18;
			this.label1.Text = "Từ ngày";
			// 
			// dtpFromDate
			// 
			this.dtpFromDate.CustomFormat = "dd/MM/yyyy HH:mm:ss";
			this.dtpFromDate.Format = System.Windows.Forms.DateTimePickerFormat.Short;
			this.dtpFromDate.Location = new System.Drawing.Point(70, 15);
			this.dtpFromDate.Name = "dtpFromDate";
			this.dtpFromDate.Size = new System.Drawing.Size(99, 20);
			this.dtpFromDate.TabIndex = 0;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(187, 18);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(53, 13);
			this.label2.TabIndex = 18;
			this.label2.Text = "Đến ngày";
			// 
			// dtpEndDate
			// 
			this.dtpEndDate.CustomFormat = "dd/MM/yyyy HH:mm:ss";
			this.dtpEndDate.Format = System.Windows.Forms.DateTimePickerFormat.Short;
			this.dtpEndDate.Location = new System.Drawing.Point(243, 15);
			this.dtpEndDate.Name = "dtpEndDate";
			this.dtpEndDate.Size = new System.Drawing.Size(98, 20);
			this.dtpEndDate.TabIndex = 1;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(360, 18);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(47, 13);
			this.label3.TabIndex = 20;
			this.label3.Text = "Từ khóa";
			// 
			// txtFindText
			// 
			this.txtFindText.Location = new System.Drawing.Point(410, 15);
			this.txtFindText.Name = "txtFindText";
			this.txtFindText.Size = new System.Drawing.Size(406, 20);
			this.txtFindText.TabIndex = 2;
			this.txtFindText.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtFindText_KeyDown);
			// 
			// btnFind
			// 
			this.btnFind.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnFind.Location = new System.Drawing.Point(858, 8);
			this.btnFind.Name = "btnFind";
			this.btnFind.Size = new System.Drawing.Size(75, 33);
			this.btnFind.TabIndex = 3;
			this.btnFind.Text = "Tìm kiếm";
			this.btnFind.UseVisualStyleBackColor = true;
			this.btnFind.Click += new System.EventHandler(this.btnFind_Click);
			// 
			// btnExportExecl
			// 
			this.btnExportExecl.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExportExecl.Appearance.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.btnExportExecl.Appearance.Options.UseFont = true;
			this.btnExportExecl.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("btnExportExecl.ImageOptions.Image")));
			this.btnExportExecl.Location = new System.Drawing.Point(899, 57);
			this.btnExportExecl.Name = "btnExportExecl";
			this.btnExportExecl.Size = new System.Drawing.Size(89, 33);
			this.btnExportExecl.TabIndex = 27;
			this.btnExportExecl.Text = "Xuất excel";
			this.btnExportExecl.Click += new System.EventHandler(this.btnExportExecl_Click);
			// 
			// sfdSaveDialog
			// 
			this.sfdSaveDialog.Title = "Lưu trữ Sound";
			// 
			// frmHistoryCheckH
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(1024, 536);
			this.Controls.Add(this.btnExportExecl);
			this.Controls.Add(this.btnFind);
			this.Controls.Add(this.txtFindText);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.dtpEndDate);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.dtpFromDate);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.grdData);
			this.Name = "frmHistoryCheckH";
			this.Text = "LỊCH SỬ KIỂM TRA SẢN PHẨM HYPONIC";
			this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
			this.Load += new System.EventHandler(this.frmHistoryCheck_Load);
			((System.ComponentModel.ISupportInitialize)(this.grdData)).EndInit();
			this.contextMenuStrip1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.grvData)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemMemoEdit1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemCheckEdit1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemCheckEdit2)).EndInit();
			this.ResumeLayout(false);
			this.PerformLayout();

        }

        #endregion
        private DevExpress.XtraGrid.GridControl grdData;
        private DevExpress.XtraGrid.Views.Grid.GridView grvData;
        private DevExpress.XtraGrid.Columns.GridColumn colPID;
        private DevExpress.XtraGrid.Columns.GridColumn colSTT;
        private DevExpress.XtraGrid.Columns.GridColumn colGroup;
        private DevExpress.XtraGrid.Columns.GridColumn colID;
        private DevExpress.XtraGrid.Columns.GridColumn colQRCode;
        private DevExpress.XtraGrid.Columns.GridColumn colOrderCode;
        private DevExpress.XtraEditors.Repository.RepositoryItemMemoEdit repositoryItemMemoEdit1;
        private DevExpress.XtraGrid.Columns.GridColumn colMonitor;
        private DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit repositoryItemCheckEdit1;
        private DevExpress.XtraGrid.Columns.GridColumn colDateLR;
        private DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit repositoryItemCheckEdit2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DateTimePicker dtpFromDate;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.DateTimePicker dtpEndDate;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtFindText;
        private System.Windows.Forms.Button btnFind;
        private DevExpress.XtraEditors.SimpleButton btnExportExecl;
		private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
		private System.Windows.Forms.ToolStripMenuItem btnShowHistoryCheckOrderDetailH;
		private System.Windows.Forms.ToolStripMenuItem nGHEÂMTHANHToolStripMenuItem;
		private System.Windows.Forms.ToolStripMenuItem nGHEÂMTHANHHYPONICTHEOQRCODERToolStripMenuItem;
		private System.Windows.Forms.SaveFileDialog sfdSaveDialog;
		private DevExpress.XtraGrid.Columns.GridColumn colMain;
		private DevExpress.XtraGrid.Columns.GridColumn colSalesOrder;
		private System.Windows.Forms.ToolStripMenuItem xEMẢNHHYPONICTHEOQRCODEToolStripMenuItem;
		private System.Windows.Forms.ToolStripMenuItem nGHELẠIFILEÂMTHANHToolStripMenuItem;
		private System.Windows.Forms.ToolStripMenuItem nGHEÂMTHANHHYPONICSẢNXUẤTLẦN1ToolStripMenuItem;
		private System.Windows.Forms.ToolStripMenuItem nGHEÂMTHANHHYPONICSỬAHÀNGKHIHÀNGNGToolStripMenuItem;
		private System.Windows.Forms.ToolStripMenuItem nGHEÂMTHANHHYPONICKIỂMTRAHÀNGKHIKHÁCHHÀNGTRẢVỀToolStripMenuItem;
		private System.Windows.Forms.ToolStripMenuItem nGHEÂMTHANHHYPONICSAUKHISỬALỖIKHÁCHHÀNGTRẢVỀToolStripMenuItem;
		private System.Windows.Forms.ToolStripMenuItem nGHEÂMTHANHCHIỀUFToolStripMenuItem;
		private System.Windows.Forms.ToolStripMenuItem nGHEÂMTHANHCHIỀUFToolStripMenuItem1;
		private System.Windows.Forms.ToolStripMenuItem nGHEÂMTHANHCHIỀUFToolStripMenuItem2;
		private System.Windows.Forms.ToolStripMenuItem nGHEÂMTHANHCHIỀUFToolStripMenuItem3;
		private System.Windows.Forms.ToolStripMenuItem nGHEÂMTHANHCHIỀURToolStripMenuItem;
		private System.Windows.Forms.ToolStripMenuItem nGHEÂMTHANHCHIỀURToolStripMenuItem1;
		private System.Windows.Forms.ToolStripMenuItem nGHEÂMTHANHCHIỀURToolStripMenuItem2;
		private System.Windows.Forms.ToolStripMenuItem nGHEÂMTHANHCHIỀURToolStripMenuItem3;
	}
}