namespace BMS
{
	partial class frmExportAndon
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmExportAndon));
            this.label1 = new System.Windows.Forms.Label();
            this.dtpFrom = new System.Windows.Forms.DateTimePicker();
            this.label2 = new System.Windows.Forms.Label();
            this.dtpTo = new System.Windows.Forms.DateTimePicker();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.txtTotalDelay = new DevExpress.XtraEditors.TextEdit();
            this.txtTotalRisk = new DevExpress.XtraEditors.TextEdit();
            this.txtTimeDelay = new DevExpress.XtraEditors.TextEdit();
            this.txtTimeRisk = new DevExpress.XtraEditors.TextEdit();
            this.btnWriteExcel = new System.Windows.Forms.Button();
            this.cbbProductStepCode = new System.Windows.Forms.ComboBox();
            this.cbbGroupProductCode = new System.Windows.Forms.ComboBox();
            this.btnGetReport = new System.Windows.Forms.Button();
            this.sfdSaveDialog = new System.Windows.Forms.SaveFileDialog();
            this.repositoryItemMemoEdit2 = new DevExpress.XtraEditors.Repository.RepositoryItemMemoEdit();
            this.grvData = new DevExpress.XtraGrid.Views.Grid.GridView();
            this.colDateTime = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colShift = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colProductCode = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colProductName = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colOrderCode = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colSerial = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colStepCode = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colWorker = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colMakeTime = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colDelayTime = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colRiskTime = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colRiskDescription = new DevExpress.XtraGrid.Columns.GridColumn();
            this.grdReport = new DevExpress.XtraGrid.GridControl();
            this.pgbExport = new System.Windows.Forms.ProgressBar();
            this.cbbType = new System.Windows.Forms.ComboBox();
            this.label9 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.txtTotalDelay.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtTotalRisk.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtTimeDelay.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtTimeRisk.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.repositoryItemMemoEdit2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.grvData)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.grdReport)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(17, 11);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(46, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Từ ngày";
            // 
            // dtpFrom
            // 
            this.dtpFrom.CustomFormat = "dd/MM/yyyy";
            this.dtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtpFrom.Location = new System.Drawing.Point(86, 8);
            this.dtpFrom.Name = "dtpFrom";
            this.dtpFrom.Size = new System.Drawing.Size(227, 20);
            this.dtpFrom.TabIndex = 1;
            this.dtpFrom.Tag = "1";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(324, 11);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(53, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "Đến ngày";
            // 
            // dtpTo
            // 
            this.dtpTo.CustomFormat = "dd/MM/yyyy";
            this.dtpTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtpTo.Location = new System.Drawing.Point(381, 8);
            this.dtpTo.Name = "dtpTo";
            this.dtpTo.Size = new System.Drawing.Size(237, 20);
            this.dtpTo.TabIndex = 2;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(17, 42);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(61, 13);
            this.label3.TabIndex = 4;
            this.label3.Text = "Công Đoạn";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(17, 77);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(56, 13);
            this.label4.TabIndex = 6;
            this.label4.Text = "Loại Hàng";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(211, 42);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(71, 13);
            this.label5.TabIndex = 8;
            this.label5.Text = "Số Lần Delay";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(210, 77);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(73, 13);
            this.label6.TabIndex = 9;
            this.label6.Text = "Số Lần Sự Cố";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(412, 42);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(83, 13);
            this.label7.TabIndex = 12;
            this.label7.Text = "Thời Gian Delay";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(412, 77);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(85, 13);
            this.label8.TabIndex = 13;
            this.label8.Text = "Thời Gian Sự Cố";
            // 
            // txtTotalDelay
            // 
            this.txtTotalDelay.Location = new System.Drawing.Point(294, 42);
            this.txtTotalDelay.Name = "txtTotalDelay";
            this.txtTotalDelay.Properties.Appearance.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
            this.txtTotalDelay.Properties.Appearance.Options.UseFont = true;
            this.txtTotalDelay.Size = new System.Drawing.Size(112, 20);
            this.txtTotalDelay.TabIndex = 5;
            // 
            // txtTotalRisk
            // 
            this.txtTotalRisk.Location = new System.Drawing.Point(294, 74);
            this.txtTotalRisk.Name = "txtTotalRisk";
            this.txtTotalRisk.Properties.Appearance.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
            this.txtTotalRisk.Properties.Appearance.Options.UseFont = true;
            this.txtTotalRisk.Size = new System.Drawing.Size(112, 20);
            this.txtTotalRisk.TabIndex = 7;
            // 
            // txtTimeDelay
            // 
            this.txtTimeDelay.EditValue = "";
            this.txtTimeDelay.Location = new System.Drawing.Point(502, 42);
            this.txtTimeDelay.Name = "txtTimeDelay";
            this.txtTimeDelay.Properties.Appearance.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
            this.txtTimeDelay.Properties.Appearance.Options.UseFont = true;
            this.txtTimeDelay.Size = new System.Drawing.Size(116, 20);
            this.txtTimeDelay.TabIndex = 6;
            // 
            // txtTimeRisk
            // 
            this.txtTimeRisk.Location = new System.Drawing.Point(502, 74);
            this.txtTimeRisk.Name = "txtTimeRisk";
            this.txtTimeRisk.Properties.Appearance.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
            this.txtTimeRisk.Properties.Appearance.Options.UseFont = true;
            this.txtTimeRisk.Size = new System.Drawing.Size(116, 20);
            this.txtTimeRisk.TabIndex = 8;
            // 
            // btnWriteExcel
            // 
            this.btnWriteExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnWriteExcel.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnWriteExcel.Location = new System.Drawing.Point(1001, 8);
            this.btnWriteExcel.Name = "btnWriteExcel";
            this.btnWriteExcel.Size = new System.Drawing.Size(170, 86);
            this.btnWriteExcel.TabIndex = 23;
            this.btnWriteExcel.Text = "Ghi File Excel";
            this.btnWriteExcel.UseVisualStyleBackColor = true;
            this.btnWriteExcel.Click += new System.EventHandler(this.btnWriteExcel_Click);
            // 
            // cbbProductStepCode
            // 
            this.cbbProductStepCode.FormattingEnabled = true;
            this.cbbProductStepCode.Location = new System.Drawing.Point(84, 39);
            this.cbbProductStepCode.Name = "cbbProductStepCode";
            this.cbbProductStepCode.Size = new System.Drawing.Size(121, 21);
            this.cbbProductStepCode.TabIndex = 3;
            // 
            // cbbGroupProductCode
            // 
            this.cbbGroupProductCode.FormattingEnabled = true;
            this.cbbGroupProductCode.Location = new System.Drawing.Point(83, 74);
            this.cbbGroupProductCode.Name = "cbbGroupProductCode";
            this.cbbGroupProductCode.Size = new System.Drawing.Size(121, 21);
            this.cbbGroupProductCode.TabIndex = 4;
            // 
            // btnGetReport
            // 
            this.btnGetReport.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnGetReport.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnGetReport.Location = new System.Drawing.Point(837, 8);
            this.btnGetReport.Name = "btnGetReport";
            this.btnGetReport.Size = new System.Drawing.Size(158, 86);
            this.btnGetReport.TabIndex = 26;
            this.btnGetReport.Text = "Xuất Báo Cáo";
            this.btnGetReport.UseVisualStyleBackColor = true;
            this.btnGetReport.Click += new System.EventHandler(this.btnGetReport_Click);
            // 
            // sfdSaveDialog
            // 
            this.sfdSaveDialog.Filter = "Excel Worksheets|*.xls";
            this.sfdSaveDialog.Title = "Lưu Trữ Báo Cáo";
            // 
            // repositoryItemMemoEdit2
            // 
            this.repositoryItemMemoEdit2.Appearance.Options.UseTextOptions = true;
            this.repositoryItemMemoEdit2.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.repositoryItemMemoEdit2.Appearance.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.repositoryItemMemoEdit2.Appearance.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
            this.repositoryItemMemoEdit2.Name = "repositoryItemMemoEdit2";
            // 
            // grvData
            // 
            this.grvData.Appearance.FocusedCell.BorderColor = System.Drawing.Color.Lime;
            this.grvData.Appearance.FocusedCell.Options.UseBorderColor = true;
            this.grvData.Appearance.HeaderPanel.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(51)))), ((int)(((byte)(204)))));
            this.grvData.Appearance.HeaderPanel.ForeColor = System.Drawing.Color.White;
            this.grvData.Appearance.HeaderPanel.Options.UseBackColor = true;
            this.grvData.Appearance.HeaderPanel.Options.UseBorderColor = true;
            this.grvData.Appearance.HeaderPanel.Options.UseFont = true;
            this.grvData.Appearance.HeaderPanel.Options.UseForeColor = true;
            this.grvData.Appearance.HeaderPanel.Options.UseImage = true;
            this.grvData.Appearance.HeaderPanel.Options.UseTextOptions = true;
            this.grvData.Appearance.HeaderPanel.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.grvData.Appearance.HeaderPanel.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.grvData.Appearance.HeaderPanel.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
            this.grvData.Appearance.ViewCaption.Options.UseTextOptions = true;
            this.grvData.Appearance.ViewCaption.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
            this.grvData.ColumnPanelRowHeight = 50;
            this.grvData.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
            this.colDateTime,
            this.colShift,
            this.colProductCode,
            this.colProductName,
            this.colOrderCode,
            this.colSerial,
            this.colStepCode,
            this.colWorker,
            this.colMakeTime,
            this.colDelayTime,
            this.colRiskTime,
            this.colRiskDescription});
            this.grvData.CustomizationFormBounds = new System.Drawing.Rectangle(1150, 183, 216, 393);
            this.grvData.GridControl = this.grdReport;
            this.grvData.HorzScrollStep = 5;
            this.grvData.Name = "grvData";
            this.grvData.OptionsBehavior.AutoExpandAllGroups = true;
            this.grvData.OptionsBehavior.Editable = false;
            this.grvData.OptionsLayout.Columns.AddNewColumns = false;
            this.grvData.OptionsLayout.Columns.RemoveOldColumns = false;
            this.grvData.OptionsSelection.EnableAppearanceFocusedCell = false;
            this.grvData.OptionsView.ColumnAutoWidth = false;
            this.grvData.OptionsView.RowAutoHeight = true;
            this.grvData.OptionsView.ShowGroupPanel = false;
            this.grvData.PaintStyleName = "Flat";
            // 
            // colDateTime
            // 
            this.colDateTime.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.colDateTime.AppearanceCell.Options.UseFont = true;
            this.colDateTime.AppearanceCell.Options.UseTextOptions = true;
            this.colDateTime.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.colDateTime.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.colDateTime.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
            this.colDateTime.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(128)))), ((int)(((byte)(255)))));
            this.colDateTime.AppearanceHeader.BackColor2 = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(128)))), ((int)(((byte)(255)))));
            this.colDateTime.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.colDateTime.AppearanceHeader.Options.UseFont = true;
            this.colDateTime.AppearanceHeader.Options.UseTextOptions = true;
            this.colDateTime.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.colDateTime.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
            this.colDateTime.Caption = "Ngày/Tháng/Năm";
            this.colDateTime.DisplayFormat.FormatString = "dd/MM/yyyy";
            this.colDateTime.DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
            this.colDateTime.FieldName = "DateReport";
            this.colDateTime.Name = "colDateTime";
            this.colDateTime.OptionsColumn.AllowEdit = false;
            this.colDateTime.OptionsColumn.AllowFocus = false;
            this.colDateTime.OptionsColumn.AllowMerge = DevExpress.Utils.DefaultBoolean.False;
            this.colDateTime.OptionsColumn.AllowMove = false;
            this.colDateTime.Visible = true;
            this.colDateTime.VisibleIndex = 0;
            this.colDateTime.Width = 139;
            // 
            // colShift
            // 
            this.colShift.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.colShift.AppearanceCell.Options.UseFont = true;
            this.colShift.AppearanceCell.Options.UseTextOptions = true;
            this.colShift.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.colShift.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.colShift.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
            this.colShift.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.colShift.AppearanceHeader.Options.UseFont = true;
            this.colShift.AppearanceHeader.Options.UseTextOptions = true;
            this.colShift.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.colShift.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
            this.colShift.Caption = "Ca/Kíp";
            this.colShift.FieldName = "ShiftName";
            this.colShift.Name = "colShift";
            this.colShift.OptionsColumn.AllowEdit = false;
            this.colShift.OptionsColumn.AllowFocus = false;
            this.colShift.OptionsColumn.AllowMerge = DevExpress.Utils.DefaultBoolean.False;
            this.colShift.OptionsColumn.AllowMove = false;
            this.colShift.Visible = true;
            this.colShift.VisibleIndex = 1;
            this.colShift.Width = 110;
            // 
            // colProductCode
            // 
            this.colProductCode.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.colProductCode.AppearanceCell.Options.UseFont = true;
            this.colProductCode.AppearanceCell.Options.UseTextOptions = true;
            this.colProductCode.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.colProductCode.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.colProductCode.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
            this.colProductCode.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.colProductCode.AppearanceHeader.Options.UseFont = true;
            this.colProductCode.Caption = "Mã Hàng";
            this.colProductCode.ColumnEdit = this.repositoryItemMemoEdit2;
            this.colProductCode.FieldName = "ProductCode";
            this.colProductCode.Name = "colProductCode";
            this.colProductCode.OptionsColumn.AllowEdit = false;
            this.colProductCode.OptionsColumn.AllowFocus = false;
            this.colProductCode.OptionsColumn.AllowMerge = DevExpress.Utils.DefaultBoolean.False;
            this.colProductCode.OptionsColumn.AllowMove = false;
            this.colProductCode.Visible = true;
            this.colProductCode.VisibleIndex = 2;
            this.colProductCode.Width = 132;
            // 
            // colProductName
            // 
            this.colProductName.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.colProductName.AppearanceCell.Options.UseFont = true;
            this.colProductName.AppearanceCell.Options.UseTextOptions = true;
            this.colProductName.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.colProductName.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.colProductName.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
            this.colProductName.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.colProductName.AppearanceHeader.Options.UseFont = true;
            this.colProductName.AppearanceHeader.Options.UseTextOptions = true;
            this.colProductName.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.colProductName.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
            this.colProductName.Caption = "Mã Mô Tả";
            this.colProductName.FieldName = "ProductName";
            this.colProductName.Name = "colProductName";
            this.colProductName.OptionsColumn.AllowEdit = false;
            this.colProductName.OptionsColumn.AllowFocus = false;
            this.colProductName.OptionsColumn.AllowMerge = DevExpress.Utils.DefaultBoolean.False;
            this.colProductName.OptionsColumn.AllowMove = false;
            this.colProductName.Visible = true;
            this.colProductName.VisibleIndex = 3;
            this.colProductName.Width = 134;
            // 
            // colOrderCode
            // 
            this.colOrderCode.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.colOrderCode.AppearanceCell.Options.UseFont = true;
            this.colOrderCode.AppearanceCell.Options.UseTextOptions = true;
            this.colOrderCode.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.colOrderCode.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.colOrderCode.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.colOrderCode.AppearanceHeader.Options.UseFont = true;
            this.colOrderCode.AppearanceHeader.Options.UseTextOptions = true;
            this.colOrderCode.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.colOrderCode.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
            this.colOrderCode.Caption = "Mã Order";
            this.colOrderCode.FieldName = "OrderCode";
            this.colOrderCode.Name = "colOrderCode";
            this.colOrderCode.OptionsColumn.AllowEdit = false;
            this.colOrderCode.OptionsColumn.AllowFocus = false;
            this.colOrderCode.OptionsColumn.AllowMerge = DevExpress.Utils.DefaultBoolean.False;
            this.colOrderCode.OptionsColumn.AllowMove = false;
            this.colOrderCode.Visible = true;
            this.colOrderCode.VisibleIndex = 4;
            this.colOrderCode.Width = 108;
            // 
            // colSerial
            // 
            this.colSerial.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.colSerial.AppearanceCell.Options.UseFont = true;
            this.colSerial.AppearanceCell.Options.UseTextOptions = true;
            this.colSerial.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.colSerial.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.colSerial.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
            this.colSerial.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.colSerial.AppearanceHeader.Options.UseFont = true;
            this.colSerial.AppearanceHeader.Options.UseTextOptions = true;
            this.colSerial.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.colSerial.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
            this.colSerial.Caption = "Số Serial";
            this.colSerial.FieldName = "QrCode";
            this.colSerial.Name = "colSerial";
            this.colSerial.OptionsColumn.AllowEdit = false;
            this.colSerial.OptionsColumn.AllowFocus = false;
            this.colSerial.OptionsColumn.AllowMerge = DevExpress.Utils.DefaultBoolean.False;
            this.colSerial.OptionsColumn.AllowMove = false;
            this.colSerial.Visible = true;
            this.colSerial.VisibleIndex = 5;
            this.colSerial.Width = 103;
            // 
            // colStepCode
            // 
            this.colStepCode.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.colStepCode.AppearanceCell.Options.UseFont = true;
            this.colStepCode.AppearanceCell.Options.UseTextOptions = true;
            this.colStepCode.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.colStepCode.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.colStepCode.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.colStepCode.AppearanceHeader.Options.UseFont = true;
            this.colStepCode.AppearanceHeader.Options.UseTextOptions = true;
            this.colStepCode.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.colStepCode.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
            this.colStepCode.Caption = "Tên Công Đoạn";
            this.colStepCode.FieldName = "ProductStepCode";
            this.colStepCode.Name = "colStepCode";
            this.colStepCode.OptionsColumn.AllowEdit = false;
            this.colStepCode.OptionsColumn.AllowFocus = false;
            this.colStepCode.OptionsColumn.AllowMerge = DevExpress.Utils.DefaultBoolean.False;
            this.colStepCode.OptionsColumn.AllowMove = false;
            this.colStepCode.Visible = true;
            this.colStepCode.VisibleIndex = 6;
            this.colStepCode.Width = 97;
            // 
            // colWorker
            // 
            this.colWorker.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.colWorker.AppearanceCell.Options.UseFont = true;
            this.colWorker.AppearanceCell.Options.UseTextOptions = true;
            this.colWorker.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.colWorker.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.colWorker.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
            this.colWorker.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.colWorker.AppearanceHeader.Options.UseFont = true;
            this.colWorker.AppearanceHeader.Options.UseTextOptions = true;
            this.colWorker.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.colWorker.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
            this.colWorker.Caption = "Người Thao Tác";
            this.colWorker.FieldName = "WorkerCode";
            this.colWorker.Name = "colWorker";
            this.colWorker.OptionsColumn.AllowEdit = false;
            this.colWorker.OptionsColumn.AllowFocus = false;
            this.colWorker.OptionsColumn.AllowMerge = DevExpress.Utils.DefaultBoolean.False;
            this.colWorker.OptionsColumn.AllowMove = false;
            this.colWorker.Visible = true;
            this.colWorker.VisibleIndex = 7;
            this.colWorker.Width = 138;
            // 
            // colMakeTime
            // 
            this.colMakeTime.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.colMakeTime.AppearanceCell.Options.UseFont = true;
            this.colMakeTime.AppearanceCell.Options.UseTextOptions = true;
            this.colMakeTime.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.colMakeTime.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.colMakeTime.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.colMakeTime.AppearanceHeader.Options.UseFont = true;
            this.colMakeTime.AppearanceHeader.Options.UseTextOptions = true;
            this.colMakeTime.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.colMakeTime.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
            this.colMakeTime.Caption = "Thời Gian Thao Tác";
            this.colMakeTime.FieldName = "MakeTime";
            this.colMakeTime.Name = "colMakeTime";
            this.colMakeTime.OptionsColumn.AllowEdit = false;
            this.colMakeTime.OptionsColumn.AllowFocus = false;
            this.colMakeTime.OptionsColumn.AllowMerge = DevExpress.Utils.DefaultBoolean.False;
            this.colMakeTime.OptionsColumn.AllowMove = false;
            this.colMakeTime.Visible = true;
            this.colMakeTime.VisibleIndex = 8;
            this.colMakeTime.Width = 107;
            // 
            // colDelayTime
            // 
            this.colDelayTime.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.colDelayTime.AppearanceCell.Options.UseFont = true;
            this.colDelayTime.AppearanceCell.Options.UseTextOptions = true;
            this.colDelayTime.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.colDelayTime.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.colDelayTime.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
            this.colDelayTime.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.colDelayTime.AppearanceHeader.Options.UseFont = true;
            this.colDelayTime.Caption = "Thời Gian Delay";
            this.colDelayTime.FieldName = "DelayTime";
            this.colDelayTime.Name = "colDelayTime";
            this.colDelayTime.OptionsColumn.AllowEdit = false;
            this.colDelayTime.OptionsColumn.AllowFocus = false;
            this.colDelayTime.OptionsColumn.AllowMerge = DevExpress.Utils.DefaultBoolean.False;
            this.colDelayTime.OptionsColumn.AllowMove = false;
            this.colDelayTime.Visible = true;
            this.colDelayTime.VisibleIndex = 9;
            this.colDelayTime.Width = 90;
            // 
            // colRiskTime
            // 
            this.colRiskTime.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.colRiskTime.AppearanceCell.Options.UseFont = true;
            this.colRiskTime.AppearanceCell.Options.UseTextOptions = true;
            this.colRiskTime.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.colRiskTime.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.colRiskTime.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
            this.colRiskTime.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.colRiskTime.AppearanceHeader.Options.UseFont = true;
            this.colRiskTime.Caption = "Thời Gian Sự Cố";
            this.colRiskTime.FieldName = "RiskTime";
            this.colRiskTime.Name = "colRiskTime";
            this.colRiskTime.OptionsColumn.AllowEdit = false;
            this.colRiskTime.OptionsColumn.AllowFocus = false;
            this.colRiskTime.OptionsColumn.AllowMerge = DevExpress.Utils.DefaultBoolean.False;
            this.colRiskTime.OptionsColumn.AllowMove = false;
            this.colRiskTime.Visible = true;
            this.colRiskTime.VisibleIndex = 10;
            // 
            // colRiskDescription
            // 
            this.colRiskDescription.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.colRiskDescription.AppearanceCell.Options.UseFont = true;
            this.colRiskDescription.AppearanceCell.Options.UseTextOptions = true;
            this.colRiskDescription.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.colRiskDescription.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.colRiskDescription.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
            this.colRiskDescription.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.colRiskDescription.AppearanceHeader.Options.UseFont = true;
            this.colRiskDescription.Caption = "Lý Do Sự Cố";
            this.colRiskDescription.FieldName = "RiskDescription";
            this.colRiskDescription.Name = "colRiskDescription";
            this.colRiskDescription.OptionsColumn.AllowEdit = false;
            this.colRiskDescription.OptionsColumn.AllowFocus = false;
            this.colRiskDescription.OptionsColumn.AllowMerge = DevExpress.Utils.DefaultBoolean.False;
            this.colRiskDescription.OptionsColumn.AllowMove = false;
            this.colRiskDescription.Visible = true;
            this.colRiskDescription.VisibleIndex = 11;
            this.colRiskDescription.Width = 132;
            // 
            // grdReport
            // 
            this.grdReport.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.grdReport.Font = new System.Drawing.Font("Times New Roman", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.grdReport.Location = new System.Drawing.Point(2, 105);
            this.grdReport.MainView = this.grvData;
            this.grdReport.Name = "grdReport";
            this.grdReport.RepositoryItems.AddRange(new DevExpress.XtraEditors.Repository.RepositoryItem[] {
            this.repositoryItemMemoEdit2});
            this.grdReport.Size = new System.Drawing.Size(1173, 512);
            this.grdReport.TabIndex = 22;
            this.grdReport.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.grvData});
            // 
            // pgbExport
            // 
            this.pgbExport.Location = new System.Drawing.Point(649, 290);
            this.pgbExport.Name = "pgbExport";
            this.pgbExport.Size = new System.Drawing.Size(213, 23);
            this.pgbExport.TabIndex = 27;
            this.pgbExport.Visible = false;
            // 
            // cbbType
            // 
            this.cbbType.FormattingEnabled = true;
            this.cbbType.Items.AddRange(new object[] {
            "Delay-Sự cố",
            "Thời gian thao tác",
            "Tất cả"});
            this.cbbType.Location = new System.Drawing.Point(633, 41);
            this.cbbType.Name = "cbbType";
            this.cbbType.Size = new System.Drawing.Size(150, 21);
            this.cbbType.TabIndex = 28;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(673, 25);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(69, 13);
            this.label9.TabIndex = 29;
            this.label9.Text = "Loại báo cáo";
            // 
            // frmExportAndon
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1174, 620);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.cbbType);
            this.Controls.Add(this.pgbExport);
            this.Controls.Add(this.btnGetReport);
            this.Controls.Add(this.cbbGroupProductCode);
            this.Controls.Add(this.cbbProductStepCode);
            this.Controls.Add(this.btnWriteExcel);
            this.Controls.Add(this.grdReport);
            this.Controls.Add(this.txtTimeRisk);
            this.Controls.Add(this.txtTimeDelay);
            this.Controls.Add(this.txtTotalRisk);
            this.Controls.Add(this.txtTotalDelay);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.dtpTo);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.dtpFrom);
            this.Controls.Add(this.label1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "frmExportAndon";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "BÁO CÁO DỮ LIỆU ANDON ALTAX";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.frmExportAndon_Load);
            ((System.ComponentModel.ISupportInitialize)(this.txtTotalDelay.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtTotalRisk.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtTimeDelay.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtTimeRisk.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.repositoryItemMemoEdit2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.grvData)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.grdReport)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

		}

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DateTimePicker dtpFrom;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.DateTimePicker dtpTo;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
        private DevExpress.XtraEditors.TextEdit txtTotalDelay;
        private DevExpress.XtraEditors.TextEdit txtTotalRisk;
        private DevExpress.XtraEditors.TextEdit txtTimeDelay;
        private DevExpress.XtraEditors.TextEdit txtTimeRisk;
        private System.Windows.Forms.Button btnWriteExcel;
        private System.Windows.Forms.ComboBox cbbProductStepCode;
        private System.Windows.Forms.ComboBox cbbGroupProductCode;
        private System.Windows.Forms.Button btnGetReport;
        private System.Windows.Forms.SaveFileDialog sfdSaveDialog;
        private DevExpress.XtraEditors.Repository.RepositoryItemMemoEdit repositoryItemMemoEdit2;
        private DevExpress.XtraGrid.Views.Grid.GridView grvData;
        private DevExpress.XtraGrid.Columns.GridColumn colDateTime;
        private DevExpress.XtraGrid.Columns.GridColumn colShift;
        private DevExpress.XtraGrid.Columns.GridColumn colProductCode;
        private DevExpress.XtraGrid.Columns.GridColumn colProductName;
        private DevExpress.XtraGrid.Columns.GridColumn colOrderCode;
        private DevExpress.XtraGrid.Columns.GridColumn colSerial;
        private DevExpress.XtraGrid.Columns.GridColumn colStepCode;
        private DevExpress.XtraGrid.Columns.GridColumn colWorker;
        private DevExpress.XtraGrid.Columns.GridColumn colMakeTime;
        private DevExpress.XtraGrid.Columns.GridColumn colDelayTime;
        private DevExpress.XtraGrid.Columns.GridColumn colRiskTime;
        private DevExpress.XtraGrid.Columns.GridColumn colRiskDescription;
        private DevExpress.XtraGrid.GridControl grdReport;
        private System.Windows.Forms.ProgressBar pgbExport;
        private System.Windows.Forms.ComboBox cbbType;
        private System.Windows.Forms.Label label9;
    }
}