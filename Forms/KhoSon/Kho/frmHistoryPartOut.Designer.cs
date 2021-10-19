
namespace BMS
{
	partial class frmHistoryPartOut
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
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmHistoryPartOut));
			this.grdData = new DevExpress.XtraGrid.GridControl();
			this.grvData = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.colID = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colStockID = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colLineName = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colPidAssembleStock = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colStockCode = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colDescriptionAssembleStock = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colStockCDCode = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colCreatDate = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colOrderCode = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colWorkerName = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colStockCDID = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colStartTime = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colEndTime = new DevExpress.XtraGrid.Columns.GridColumn();
			this.btnSearchHistory = new DevExpress.XtraEditors.SimpleButton();
			this.txbSearchHistory = new System.Windows.Forms.TextBox();
			this.label3 = new System.Windows.Forms.Label();
			this.dtpFrom = new System.Windows.Forms.DateTimePicker();
			this.panel2 = new System.Windows.Forms.Panel();
			this.btnXuatEx = new System.Windows.Forms.Button();
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
            this.colStockID,
            this.colLineName,
            this.colPidAssembleStock,
            this.colStockCode,
            this.colDescriptionAssembleStock,
            this.colStockCDCode,
            this.colCreatDate,
            this.colOrderCode,
            this.colWorkerName,
            this.colStockCDID,
            this.colStartTime,
            this.colEndTime});
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
			this.grvData.DoubleClick += new System.EventHandler(this.gvHistory_DoubleClick);
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
			// colStockID
			// 
			this.colStockID.AppearanceCell.Options.UseTextOptions = true;
			this.colStockID.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colStockID.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colStockID.AppearanceHeader.ForeColor = System.Drawing.Color.Black;
			this.colStockID.AppearanceHeader.Options.UseForeColor = true;
			this.colStockID.Caption = "StockID";
			this.colStockID.FieldName = "StockID";
			this.colStockID.Name = "colStockID";
			// 
			// colLineName
			// 
			this.colLineName.AppearanceCell.BackColor = System.Drawing.Color.White;
			this.colLineName.AppearanceCell.BackColor2 = System.Drawing.Color.White;
			this.colLineName.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.colLineName.AppearanceCell.Options.UseBackColor = true;
			this.colLineName.AppearanceCell.Options.UseFont = true;
			this.colLineName.AppearanceCell.Options.UseTextOptions = true;
			this.colLineName.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colLineName.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colLineName.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colLineName.AppearanceHeader.ForeColor = System.Drawing.Color.Black;
			this.colLineName.AppearanceHeader.Options.UseFont = true;
			this.colLineName.AppearanceHeader.Options.UseForeColor = true;
			this.colLineName.AppearanceHeader.Options.UseTextOptions = true;
			this.colLineName.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colLineName.Caption = "Line";
			this.colLineName.FieldName = "LineName";
			this.colLineName.Name = "colLineName";
			this.colLineName.OptionsColumn.AllowEdit = false;
			this.colLineName.Visible = true;
			this.colLineName.VisibleIndex = 2;
			this.colLineName.Width = 74;
			// 
			// colPidAssembleStock
			// 
			this.colPidAssembleStock.AppearanceCell.BackColor = System.Drawing.Color.White;
			this.colPidAssembleStock.AppearanceCell.BackColor2 = System.Drawing.Color.White;
			this.colPidAssembleStock.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.colPidAssembleStock.AppearanceCell.Options.UseBackColor = true;
			this.colPidAssembleStock.AppearanceCell.Options.UseFont = true;
			this.colPidAssembleStock.AppearanceCell.Options.UseTextOptions = true;
			this.colPidAssembleStock.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colPidAssembleStock.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colPidAssembleStock.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colPidAssembleStock.AppearanceHeader.ForeColor = System.Drawing.Color.Black;
			this.colPidAssembleStock.AppearanceHeader.Options.UseFont = true;
			this.colPidAssembleStock.AppearanceHeader.Options.UseForeColor = true;
			this.colPidAssembleStock.AppearanceHeader.Options.UseTextOptions = true;
			this.colPidAssembleStock.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colPidAssembleStock.Caption = "PID";
			this.colPidAssembleStock.FieldName = "PidAssembleStock";
			this.colPidAssembleStock.Name = "colPidAssembleStock";
			this.colPidAssembleStock.OptionsColumn.AllowEdit = false;
			this.colPidAssembleStock.Visible = true;
			this.colPidAssembleStock.VisibleIndex = 4;
			this.colPidAssembleStock.Width = 148;
			// 
			// colStockCode
			// 
			this.colStockCode.AppearanceCell.BackColor = System.Drawing.Color.White;
			this.colStockCode.AppearanceCell.BackColor2 = System.Drawing.Color.White;
			this.colStockCode.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.colStockCode.AppearanceCell.Options.UseBackColor = true;
			this.colStockCode.AppearanceCell.Options.UseFont = true;
			this.colStockCode.AppearanceCell.Options.UseTextOptions = true;
			this.colStockCode.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colStockCode.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colStockCode.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colStockCode.AppearanceHeader.ForeColor = System.Drawing.Color.Black;
			this.colStockCode.AppearanceHeader.Options.UseFont = true;
			this.colStockCode.AppearanceHeader.Options.UseForeColor = true;
			this.colStockCode.AppearanceHeader.Options.UseTextOptions = true;
			this.colStockCode.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colStockCode.Caption = "Tên kho";
			this.colStockCode.FieldName = "StockCode";
			this.colStockCode.Name = "colStockCode";
			this.colStockCode.OptionsColumn.AllowEdit = false;
			this.colStockCode.Visible = true;
			this.colStockCode.VisibleIndex = 0;
			this.colStockCode.Width = 89;
			// 
			// colDescriptionAssembleStock
			// 
			this.colDescriptionAssembleStock.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.colDescriptionAssembleStock.AppearanceCell.Options.UseFont = true;
			this.colDescriptionAssembleStock.AppearanceCell.Options.UseTextOptions = true;
			this.colDescriptionAssembleStock.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colDescriptionAssembleStock.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colDescriptionAssembleStock.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colDescriptionAssembleStock.AppearanceHeader.ForeColor = System.Drawing.Color.Black;
			this.colDescriptionAssembleStock.AppearanceHeader.Options.UseFont = true;
			this.colDescriptionAssembleStock.AppearanceHeader.Options.UseForeColor = true;
			this.colDescriptionAssembleStock.AppearanceHeader.Options.UseTextOptions = true;
			this.colDescriptionAssembleStock.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colDescriptionAssembleStock.Caption = "Mô tả";
			this.colDescriptionAssembleStock.FieldName = "DescriptionAssembleStock";
			this.colDescriptionAssembleStock.Name = "colDescriptionAssembleStock";
			this.colDescriptionAssembleStock.OptionsColumn.AllowEdit = false;
			this.colDescriptionAssembleStock.Visible = true;
			this.colDescriptionAssembleStock.VisibleIndex = 5;
			this.colDescriptionAssembleStock.Width = 205;
			// 
			// colStockCDCode
			// 
			this.colStockCDCode.AppearanceCell.BackColor = System.Drawing.Color.White;
			this.colStockCDCode.AppearanceCell.BackColor2 = System.Drawing.Color.White;
			this.colStockCDCode.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.colStockCDCode.AppearanceCell.Options.UseBackColor = true;
			this.colStockCDCode.AppearanceCell.Options.UseFont = true;
			this.colStockCDCode.AppearanceCell.Options.UseTextOptions = true;
			this.colStockCDCode.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colStockCDCode.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colStockCDCode.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colStockCDCode.AppearanceHeader.ForeColor = System.Drawing.Color.Black;
			this.colStockCDCode.AppearanceHeader.Options.UseFont = true;
			this.colStockCDCode.AppearanceHeader.Options.UseForeColor = true;
			this.colStockCDCode.AppearanceHeader.Options.UseTextOptions = true;
			this.colStockCDCode.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colStockCDCode.Caption = "Tên công đoạn";
			this.colStockCDCode.FieldName = "StockCDCode";
			this.colStockCDCode.Name = "colStockCDCode";
			this.colStockCDCode.OptionsColumn.AllowEdit = false;
			this.colStockCDCode.Visible = true;
			this.colStockCDCode.VisibleIndex = 1;
			this.colStockCDCode.Width = 133;
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
			// colOrderCode
			// 
			this.colOrderCode.AppearanceCell.BackColor = System.Drawing.Color.White;
			this.colOrderCode.AppearanceCell.BackColor2 = System.Drawing.Color.White;
			this.colOrderCode.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.colOrderCode.AppearanceCell.Options.UseBackColor = true;
			this.colOrderCode.AppearanceCell.Options.UseFont = true;
			this.colOrderCode.AppearanceCell.Options.UseTextOptions = true;
			this.colOrderCode.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colOrderCode.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colOrderCode.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colOrderCode.AppearanceHeader.ForeColor = System.Drawing.Color.Black;
			this.colOrderCode.AppearanceHeader.Options.UseFont = true;
			this.colOrderCode.AppearanceHeader.Options.UseForeColor = true;
			this.colOrderCode.AppearanceHeader.Options.UseTextOptions = true;
			this.colOrderCode.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colOrderCode.Caption = "Order";
			this.colOrderCode.FieldName = "OrderCode";
			this.colOrderCode.Name = "colOrderCode";
			this.colOrderCode.OptionsColumn.AllowEdit = false;
			this.colOrderCode.Visible = true;
			this.colOrderCode.VisibleIndex = 3;
			this.colOrderCode.Width = 155;
			// 
			// colWorkerName
			// 
			this.colWorkerName.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.colWorkerName.AppearanceCell.Options.UseFont = true;
			this.colWorkerName.AppearanceCell.Options.UseTextOptions = true;
			this.colWorkerName.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colWorkerName.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colWorkerName.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colWorkerName.AppearanceHeader.ForeColor = System.Drawing.Color.Black;
			this.colWorkerName.AppearanceHeader.Options.UseFont = true;
			this.colWorkerName.AppearanceHeader.Options.UseForeColor = true;
			this.colWorkerName.AppearanceHeader.Options.UseTextOptions = true;
			this.colWorkerName.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colWorkerName.Caption = "Người nhập";
			this.colWorkerName.FieldName = "WorkerName";
			this.colWorkerName.Name = "colWorkerName";
			this.colWorkerName.Visible = true;
			this.colWorkerName.VisibleIndex = 6;
			this.colWorkerName.Width = 129;
			// 
			// colStockCDID
			// 
			this.colStockCDID.AppearanceHeader.ForeColor = System.Drawing.Color.Black;
			this.colStockCDID.AppearanceHeader.Options.UseForeColor = true;
			this.colStockCDID.Caption = "StockCDID";
			this.colStockCDID.FieldName = "StockCDID";
			this.colStockCDID.Name = "colStockCDID";
			// 
			// colStartTime
			// 
			this.colStartTime.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.colStartTime.AppearanceCell.Options.UseFont = true;
			this.colStartTime.AppearanceCell.Options.UseTextOptions = true;
			this.colStartTime.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colStartTime.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colStartTime.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colStartTime.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colStartTime.AppearanceHeader.Options.UseFont = true;
			this.colStartTime.Caption = "Thời gian bắt đầu";
			this.colStartTime.DisplayFormat.FormatString = "HH:mm:ss";
			this.colStartTime.DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
			this.colStartTime.FieldName = "StartTime";
			this.colStartTime.Name = "colStartTime";
			this.colStartTime.Visible = true;
			this.colStartTime.VisibleIndex = 8;
			this.colStartTime.Width = 172;
			// 
			// colEndTime
			// 
			this.colEndTime.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.colEndTime.AppearanceCell.Options.UseFont = true;
			this.colEndTime.AppearanceCell.Options.UseTextOptions = true;
			this.colEndTime.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colEndTime.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colEndTime.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colEndTime.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colEndTime.AppearanceHeader.Options.UseFont = true;
			this.colEndTime.Caption = "Thời gian kết thúc";
			this.colEndTime.DisplayFormat.FormatString = "HH:mm:ss";
			this.colEndTime.DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
			this.colEndTime.FieldName = "EndTime";
			this.colEndTime.Name = "colEndTime";
			this.colEndTime.Visible = true;
			this.colEndTime.VisibleIndex = 9;
			this.colEndTime.Width = 176;
			// 
			// btnSearchHistory
			// 
			this.btnSearchHistory.Appearance.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnSearchHistory.Appearance.Options.UseFont = true;
			this.btnSearchHistory.Location = new System.Drawing.Point(547, 11);
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
			this.panel2.Controls.Add(this.btnXuatEx);
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
			// btnXuatEx
			// 
			this.btnXuatEx.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.8F);
			this.btnXuatEx.Location = new System.Drawing.Point(653, 10);
			this.btnXuatEx.Name = "btnXuatEx";
			this.btnXuatEx.Size = new System.Drawing.Size(98, 27);
			this.btnXuatEx.TabIndex = 37;
			this.btnXuatEx.Text = "Xuất excel";
			this.btnXuatEx.UseVisualStyleBackColor = true;
			this.btnXuatEx.Click += new System.EventHandler(this.btnXuatEx_Click);
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
			// frmHistoryPartOut
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(994, 571);
			this.Controls.Add(this.toolStrip1);
			this.Controls.Add(this.grdData);
			this.Controls.Add(this.panel2);
			this.Margin = new System.Windows.Forms.Padding(4);
			this.Name = "frmHistoryPartOut";
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
		private DevExpress.XtraGrid.Columns.GridColumn colStockCode;
		private DevExpress.XtraGrid.Columns.GridColumn colStockCDCode;
		private DevExpress.XtraGrid.Columns.GridColumn colOrderCode;
		private DevExpress.XtraGrid.Columns.GridColumn colLineName;
		private DevExpress.XtraGrid.Columns.GridColumn colPidAssembleStock;
		private DevExpress.XtraEditors.SimpleButton btnSearchHistory;
		private System.Windows.Forms.TextBox txbSearchHistory;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.DateTimePicker dtpFrom;
		private System.Windows.Forms.Panel panel2;
		private System.Windows.Forms.DateTimePicker dtpTo;
		private System.Windows.Forms.Label label1;
		private DevExpress.XtraGrid.Columns.GridColumn colStockID;
		private DevExpress.XtraGrid.Columns.GridColumn colDescriptionAssembleStock;
		private DevExpress.XtraGrid.Columns.GridColumn colCreatDate;
		private System.Windows.Forms.ToolStrip toolStrip1;
		private System.Windows.Forms.ToolStripButton btnEditPart;
		private System.Windows.Forms.ToolStripSeparator toolStripSeparator7;
		private System.Windows.Forms.ToolStripButton btnDelPart;
		private System.Windows.Forms.ToolStripSeparator toolStripSeparator8;
		private DevExpress.XtraGrid.Columns.GridColumn colWorkerName;
		private DevExpress.XtraGrid.Columns.GridColumn colStockCDID;
		private System.Windows.Forms.Button btnXuatEx;
		private DevExpress.XtraGrid.Columns.GridColumn colStartTime;
		private DevExpress.XtraGrid.Columns.GridColumn colEndTime;
	}
}