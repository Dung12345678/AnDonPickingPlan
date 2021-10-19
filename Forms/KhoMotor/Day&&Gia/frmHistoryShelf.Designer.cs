
namespace BMS
{
	partial class frmHistoryShelf
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
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmHistoryShelf));
			this.grdData = new DevExpress.XtraGrid.GridControl();
			this.grvData = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.colID = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colSerialNumber = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colQtySale = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colWorker = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colCreateDate = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colShelf = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colCreatDate = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colSaleNumBer = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colArticleID = new DevExpress.XtraGrid.Columns.GridColumn();
			this.btnSearchHistory = new DevExpress.XtraEditors.SimpleButton();
			this.txbSearchHistory = new System.Windows.Forms.TextBox();
			this.label3 = new System.Windows.Forms.Label();
			this.dtpFrom = new System.Windows.Forms.DateTimePicker();
			this.panel2 = new System.Windows.Forms.Panel();
			this.dtpTo = new System.Windows.Forms.DateTimePicker();
			this.label1 = new System.Windows.Forms.Label();
			this.toolStrip1 = new System.Windows.Forms.ToolStrip();
			this.btnEditPart = new System.Windows.Forms.ToolStripButton();
			this.toolStripSeparator7 = new System.Windows.Forms.ToolStripSeparator();
			this.btnDelPart = new System.Windows.Forms.ToolStripButton();
			this.toolStripSeparator8 = new System.Windows.Forms.ToolStripSeparator();
			((System.ComponentModel.ISupportInitialize)(this.grdData)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grvData)).BeginInit();
			this.panel2.SuspendLayout();
			this.toolStrip1.SuspendLayout();
			this.SuspendLayout();
			// 
			// grdData
			// 
			this.grdData.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.grdData.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.grdData.Location = new System.Drawing.Point(0, 45);
			this.grdData.MainView = this.grvData;
			this.grdData.Name = "grdData";
			this.grdData.Size = new System.Drawing.Size(992, 526);
			this.grdData.TabIndex = 29;
			this.grdData.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.grvData});
			// 
			// grvData
			// 
			this.grvData.ColumnPanelRowHeight = 40;
			this.grvData.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
            this.colID,
            this.colSerialNumber,
            this.colQtySale,
            this.colWorker,
            this.colCreateDate,
            this.colShelf,
            this.colCreatDate,
            this.colSaleNumBer,
            this.colArticleID});
			this.grvData.GridControl = this.grdData;
			this.grvData.Name = "grvData";
			this.grvData.OptionsBehavior.AllowAddRows = DevExpress.Utils.DefaultBoolean.False;
			this.grvData.OptionsBehavior.AllowDeleteRows = DevExpress.Utils.DefaultBoolean.False;
			this.grvData.OptionsBehavior.AllowFixedGroups = DevExpress.Utils.DefaultBoolean.False;
			this.grvData.OptionsBehavior.AllowIncrementalSearch = true;
			this.grvData.OptionsBehavior.Editable = false;
			this.grvData.OptionsView.ColumnAutoWidth = false;
			this.grvData.OptionsView.ShowGroupPanel = false;
			this.grvData.RowHeight = 25;
			// 
			// colID
			// 
			this.colID.AppearanceCell.Options.UseTextOptions = true;
			this.colID.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colID.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colID.AppearanceHeader.ForeColor = System.Drawing.Color.Black;
			this.colID.AppearanceHeader.Options.UseForeColor = true;
			this.colID.Caption = "ID";
			this.colID.FieldName = "ID";
			this.colID.Name = "colID";
			this.colID.OptionsColumn.AllowEdit = false;
			// 
			// colSerialNumber
			// 
			this.colSerialNumber.AppearanceCell.BackColor = System.Drawing.Color.White;
			this.colSerialNumber.AppearanceCell.BackColor2 = System.Drawing.Color.White;
			this.colSerialNumber.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.colSerialNumber.AppearanceCell.Options.UseBackColor = true;
			this.colSerialNumber.AppearanceCell.Options.UseFont = true;
			this.colSerialNumber.AppearanceCell.Options.UseTextOptions = true;
			this.colSerialNumber.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colSerialNumber.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colSerialNumber.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colSerialNumber.AppearanceHeader.ForeColor = System.Drawing.Color.Black;
			this.colSerialNumber.AppearanceHeader.Options.UseFont = true;
			this.colSerialNumber.AppearanceHeader.Options.UseForeColor = true;
			this.colSerialNumber.AppearanceHeader.Options.UseTextOptions = true;
			this.colSerialNumber.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colSerialNumber.Caption = "Serial Number";
			this.colSerialNumber.FieldName = "SerialNumber";
			this.colSerialNumber.Name = "colSerialNumber";
			this.colSerialNumber.OptionsColumn.AllowEdit = false;
			this.colSerialNumber.Visible = true;
			this.colSerialNumber.VisibleIndex = 2;
			this.colSerialNumber.Width = 100;
			// 
			// colQtySale
			// 
			this.colQtySale.AppearanceCell.BackColor = System.Drawing.Color.White;
			this.colQtySale.AppearanceCell.BackColor2 = System.Drawing.Color.White;
			this.colQtySale.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.colQtySale.AppearanceCell.Options.UseBackColor = true;
			this.colQtySale.AppearanceCell.Options.UseFont = true;
			this.colQtySale.AppearanceCell.Options.UseTextOptions = true;
			this.colQtySale.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colQtySale.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colQtySale.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colQtySale.AppearanceHeader.ForeColor = System.Drawing.Color.Black;
			this.colQtySale.AppearanceHeader.Options.UseFont = true;
			this.colQtySale.AppearanceHeader.Options.UseForeColor = true;
			this.colQtySale.AppearanceHeader.Options.UseTextOptions = true;
			this.colQtySale.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colQtySale.Caption = "QtySale";
			this.colQtySale.FieldName = "QtySale";
			this.colQtySale.Name = "colQtySale";
			this.colQtySale.OptionsColumn.AllowEdit = false;
			this.colQtySale.Visible = true;
			this.colQtySale.VisibleIndex = 4;
			this.colQtySale.Width = 148;
			// 
			// colWorker
			// 
			this.colWorker.AppearanceCell.BackColor = System.Drawing.Color.White;
			this.colWorker.AppearanceCell.BackColor2 = System.Drawing.Color.White;
			this.colWorker.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.colWorker.AppearanceCell.Options.UseBackColor = true;
			this.colWorker.AppearanceCell.Options.UseFont = true;
			this.colWorker.AppearanceCell.Options.UseTextOptions = true;
			this.colWorker.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colWorker.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colWorker.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colWorker.AppearanceHeader.ForeColor = System.Drawing.Color.Black;
			this.colWorker.AppearanceHeader.Options.UseFont = true;
			this.colWorker.AppearanceHeader.Options.UseForeColor = true;
			this.colWorker.AppearanceHeader.Options.UseTextOptions = true;
			this.colWorker.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colWorker.Caption = "Tên người nhập";
			this.colWorker.FieldName = "Worker";
			this.colWorker.Name = "colWorker";
			this.colWorker.OptionsColumn.AllowEdit = false;
			this.colWorker.Visible = true;
			this.colWorker.VisibleIndex = 0;
			this.colWorker.Width = 129;
			// 
			// colCreateDate
			// 
			this.colCreateDate.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.colCreateDate.AppearanceCell.Options.UseFont = true;
			this.colCreateDate.AppearanceCell.Options.UseTextOptions = true;
			this.colCreateDate.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colCreateDate.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colCreateDate.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colCreateDate.AppearanceHeader.ForeColor = System.Drawing.Color.Black;
			this.colCreateDate.AppearanceHeader.Options.UseFont = true;
			this.colCreateDate.AppearanceHeader.Options.UseForeColor = true;
			this.colCreateDate.AppearanceHeader.Options.UseTextOptions = true;
			this.colCreateDate.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colCreateDate.Caption = "Create Date";
			this.colCreateDate.FieldName = "CreateDate";
			this.colCreateDate.Name = "colCreateDate";
			this.colCreateDate.OptionsColumn.AllowEdit = false;
			this.colCreateDate.Visible = true;
			this.colCreateDate.VisibleIndex = 6;
			this.colCreateDate.Width = 205;
			// 
			// colShelf
			// 
			this.colShelf.AppearanceCell.BackColor = System.Drawing.Color.White;
			this.colShelf.AppearanceCell.BackColor2 = System.Drawing.Color.White;
			this.colShelf.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.colShelf.AppearanceCell.Options.UseBackColor = true;
			this.colShelf.AppearanceCell.Options.UseFont = true;
			this.colShelf.AppearanceCell.Options.UseTextOptions = true;
			this.colShelf.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colShelf.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colShelf.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colShelf.AppearanceHeader.ForeColor = System.Drawing.Color.Black;
			this.colShelf.AppearanceHeader.Options.UseFont = true;
			this.colShelf.AppearanceHeader.Options.UseForeColor = true;
			this.colShelf.AppearanceHeader.Options.UseTextOptions = true;
			this.colShelf.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colShelf.Caption = "Giá đỡ";
			this.colShelf.FieldName = "Shelf";
			this.colShelf.Name = "colShelf";
			this.colShelf.OptionsColumn.AllowEdit = false;
			this.colShelf.Visible = true;
			this.colShelf.VisibleIndex = 1;
			this.colShelf.Width = 133;
			// 
			// colCreatDate
			// 
			this.colCreatDate.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.colCreatDate.AppearanceCell.Options.UseFont = true;
			this.colCreatDate.AppearanceCell.Options.UseTextOptions = true;
			this.colCreatDate.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colCreatDate.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colCreatDate.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colCreatDate.AppearanceHeader.ForeColor = System.Drawing.Color.Black;
			this.colCreatDate.AppearanceHeader.Options.UseFont = true;
			this.colCreatDate.AppearanceHeader.Options.UseForeColor = true;
			this.colCreatDate.AppearanceHeader.Options.UseTextOptions = true;
			this.colCreatDate.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colCreatDate.Caption = "Thời gian";
			this.colCreatDate.DisplayFormat.FormatString = "MM/dd/yyyy HH:mm";
			this.colCreatDate.DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
			this.colCreatDate.FieldName = "CreatDate";
			this.colCreatDate.GroupFormat.FormatString = "d";
			this.colCreatDate.GroupFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
			this.colCreatDate.Name = "colCreatDate";
			this.colCreatDate.OptionsColumn.AllowEdit = false;
			this.colCreatDate.Visible = true;
			this.colCreatDate.VisibleIndex = 7;
			this.colCreatDate.Width = 179;
			// 
			// colSaleNumBer
			// 
			this.colSaleNumBer.AppearanceCell.BackColor = System.Drawing.Color.White;
			this.colSaleNumBer.AppearanceCell.BackColor2 = System.Drawing.Color.White;
			this.colSaleNumBer.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.colSaleNumBer.AppearanceCell.Options.UseBackColor = true;
			this.colSaleNumBer.AppearanceCell.Options.UseFont = true;
			this.colSaleNumBer.AppearanceCell.Options.UseTextOptions = true;
			this.colSaleNumBer.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colSaleNumBer.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colSaleNumBer.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colSaleNumBer.AppearanceHeader.ForeColor = System.Drawing.Color.Black;
			this.colSaleNumBer.AppearanceHeader.Options.UseFont = true;
			this.colSaleNumBer.AppearanceHeader.Options.UseForeColor = true;
			this.colSaleNumBer.AppearanceHeader.Options.UseTextOptions = true;
			this.colSaleNumBer.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colSaleNumBer.Caption = "Sale NumBer";
			this.colSaleNumBer.FieldName = "SaleNumBer";
			this.colSaleNumBer.Name = "colSaleNumBer";
			this.colSaleNumBer.OptionsColumn.AllowEdit = false;
			this.colSaleNumBer.Visible = true;
			this.colSaleNumBer.VisibleIndex = 3;
			this.colSaleNumBer.Width = 155;
			// 
			// colArticleID
			// 
			this.colArticleID.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.colArticleID.AppearanceCell.Options.UseFont = true;
			this.colArticleID.AppearanceCell.Options.UseTextOptions = true;
			this.colArticleID.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colArticleID.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colArticleID.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colArticleID.AppearanceHeader.ForeColor = System.Drawing.Color.Black;
			this.colArticleID.AppearanceHeader.Options.UseFont = true;
			this.colArticleID.AppearanceHeader.Options.UseForeColor = true;
			this.colArticleID.AppearanceHeader.Options.UseTextOptions = true;
			this.colArticleID.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colArticleID.Caption = "ArticleID";
			this.colArticleID.FieldName = "ArticleID";
			this.colArticleID.Name = "colArticleID";
			this.colArticleID.Visible = true;
			this.colArticleID.VisibleIndex = 5;
			this.colArticleID.Width = 129;
			// 
			// btnSearchHistory
			// 
			this.btnSearchHistory.Appearance.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnSearchHistory.Appearance.Options.UseFont = true;
			this.btnSearchHistory.Location = new System.Drawing.Point(547, 10);
			this.btnSearchHistory.Name = "btnSearchHistory";
			this.btnSearchHistory.Size = new System.Drawing.Size(75, 23);
			this.btnSearchHistory.TabIndex = 25;
			this.btnSearchHistory.Text = "Tìm kiếm";
			this.btnSearchHistory.Click += new System.EventHandler(this.btnSearchHistory_Click);
			// 
			// txbSearchHistory
			// 
			this.txbSearchHistory.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txbSearchHistory.Location = new System.Drawing.Point(394, 10);
			this.txbSearchHistory.Name = "txbSearchHistory";
			this.txbSearchHistory.Size = new System.Drawing.Size(147, 24);
			this.txbSearchHistory.TabIndex = 24;
			this.txbSearchHistory.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txbSearchHistory_KeyDown);
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.label3.Location = new System.Drawing.Point(12, 11);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(68, 18);
			this.label3.TabIndex = 33;
			this.label3.Text = "Từ ngày: ";
			// 
			// dtpFrom
			// 
			this.dtpFrom.CustomFormat = "dd/MM/yyyy";
			this.dtpFrom.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.dtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
			this.dtpFrom.Location = new System.Drawing.Point(74, 10);
			this.dtpFrom.Name = "dtpFrom";
			this.dtpFrom.Size = new System.Drawing.Size(106, 24);
			this.dtpFrom.TabIndex = 36;
			this.dtpFrom.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtpFrom_KeyPress);
			// 
			// panel2
			// 
			this.panel2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.panel2.Controls.Add(this.btnSearchHistory);
			this.panel2.Controls.Add(this.txbSearchHistory);
			this.panel2.Controls.Add(this.dtpFrom);
			this.panel2.Controls.Add(this.dtpTo);
			this.panel2.Controls.Add(this.label1);
			this.panel2.Controls.Add(this.label3);
			this.panel2.Location = new System.Drawing.Point(0, 1);
			this.panel2.Name = "panel2";
			this.panel2.Size = new System.Drawing.Size(992, 42);
			this.panel2.TabIndex = 4;
			// 
			// dtpTo
			// 
			this.dtpTo.CustomFormat = "dd/MM/yyyy";
			this.dtpTo.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.dtpTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
			this.dtpTo.Location = new System.Drawing.Point(274, 10);
			this.dtpTo.Name = "dtpTo";
			this.dtpTo.Size = new System.Drawing.Size(99, 24);
			this.dtpTo.TabIndex = 35;
			this.dtpTo.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtpTo_KeyPress);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.label1.Location = new System.Drawing.Point(205, 11);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(78, 18);
			this.label1.TabIndex = 34;
			this.label1.Text = "Đến ngày: ";
			// 
			// toolStrip1
			// 
			this.toolStrip1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.toolStrip1.AutoSize = false;
			this.toolStrip1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(234)))), ((int)(((byte)(234)))), ((int)(((byte)(234)))));
			this.toolStrip1.Dock = System.Windows.Forms.DockStyle.None;
			this.toolStrip1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.toolStrip1.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden;
			this.toolStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
			this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.btnEditPart,
            this.toolStripSeparator7,
            this.btnDelPart,
            this.toolStripSeparator8});
			this.toolStrip1.LayoutStyle = System.Windows.Forms.ToolStripLayoutStyle.Flow;
			this.toolStrip1.Location = new System.Drawing.Point(0, 41);
			this.toolStrip1.Name = "toolStrip1";
			this.toolStrip1.RenderMode = System.Windows.Forms.ToolStripRenderMode.Professional;
			this.toolStrip1.Size = new System.Drawing.Size(994, 39);
			this.toolStrip1.TabIndex = 30;
			this.toolStrip1.Text = "toolStrip2";
			this.toolStrip1.Visible = false;
			// 
			// btnEditPart
			// 
			this.btnEditPart.Enabled = false;
			this.btnEditPart.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnEditPart.Image = ((System.Drawing.Image)(resources.GetObject("btnEditPart.Image")));
			this.btnEditPart.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnEditPart.Name = "btnEditPart";
			this.btnEditPart.Size = new System.Drawing.Size(72, 37);
			this.btnEditPart.Tag = "frmProduct_EditProductH";
			this.btnEditPart.Text = "Sửa lịch sử";
			this.btnEditPart.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			// 
			// toolStripSeparator7
			// 
			this.toolStripSeparator7.Name = "toolStripSeparator7";
			this.toolStripSeparator7.Size = new System.Drawing.Size(6, 23);
			// 
			// btnDelPart
			// 
			this.btnDelPart.Enabled = false;
			this.btnDelPart.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnDelPart.Image = ((System.Drawing.Image)(resources.GetObject("btnDelPart.Image")));
			this.btnDelPart.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnDelPart.Name = "btnDelPart";
			this.btnDelPart.Size = new System.Drawing.Size(71, 37);
			this.btnDelPart.Tag = "frmProduct_DeleteProductH";
			this.btnDelPart.Text = "Xóa lịch sử";
			this.btnDelPart.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			// 
			// toolStripSeparator8
			// 
			this.toolStripSeparator8.AutoSize = false;
			this.toolStripSeparator8.Name = "toolStripSeparator8";
			this.toolStripSeparator8.Size = new System.Drawing.Size(6, 40);
			// 
			// frmHistoryShelf
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(994, 571);
			this.Controls.Add(this.toolStrip1);
			this.Controls.Add(this.grdData);
			this.Controls.Add(this.panel2);
			this.Margin = new System.Windows.Forms.Padding(4);
			this.Name = "frmHistoryShelf";
			this.Text = "LỊCH SỬ XUẤT KHO LẮP RÁP";
			this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
			this.Load += new System.EventHandler(this.frmHistoryPartOut_Load);
			((System.ComponentModel.ISupportInitialize)(this.grdData)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grvData)).EndInit();
			this.panel2.ResumeLayout(false);
			this.panel2.PerformLayout();
			this.toolStrip1.ResumeLayout(false);
			this.toolStrip1.PerformLayout();
			this.ResumeLayout(false);

		}

		#endregion
		private DevExpress.XtraGrid.GridControl grdData;
		private DevExpress.XtraGrid.Views.Grid.GridView grvData;
		private DevExpress.XtraGrid.Columns.GridColumn colID;
		private DevExpress.XtraGrid.Columns.GridColumn colWorker;
		private DevExpress.XtraGrid.Columns.GridColumn colShelf;
		private DevExpress.XtraGrid.Columns.GridColumn colSaleNumBer;
		private DevExpress.XtraGrid.Columns.GridColumn colSerialNumber;
		private DevExpress.XtraGrid.Columns.GridColumn colQtySale;
		private DevExpress.XtraEditors.SimpleButton btnSearchHistory;
		private System.Windows.Forms.TextBox txbSearchHistory;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.DateTimePicker dtpFrom;
		private System.Windows.Forms.Panel panel2;
		private System.Windows.Forms.DateTimePicker dtpTo;
		private System.Windows.Forms.Label label1;
		private DevExpress.XtraGrid.Columns.GridColumn colCreateDate;
		private DevExpress.XtraGrid.Columns.GridColumn colCreatDate;
		private System.Windows.Forms.ToolStrip toolStrip1;
		private System.Windows.Forms.ToolStripButton btnEditPart;
		private System.Windows.Forms.ToolStripSeparator toolStripSeparator7;
		private System.Windows.Forms.ToolStripButton btnDelPart;
		private System.Windows.Forms.ToolStripSeparator toolStripSeparator8;
		private DevExpress.XtraGrid.Columns.GridColumn colArticleID;
	}
}