namespace BMS
{
    partial class frmReportTester
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
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmReportTester));
			this.label1 = new System.Windows.Forms.Label();
			this.dtpFrom = new System.Windows.Forms.DateTimePicker();
			this.label2 = new System.Windows.Forms.Label();
			this.dtpTo = new System.Windows.Forms.DateTimePicker();
			this.label3 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.txtLot = new DevExpress.XtraEditors.TextEdit();
			this.btnWriteExcel = new System.Windows.Forms.Button();
			this.btnGetReport = new System.Windows.Forms.Button();
			this.sfdSaveDialog = new System.Windows.Forms.SaveFileDialog();
			this.repositoryItemMemoEdit2 = new DevExpress.XtraEditors.Repository.RepositoryItemMemoEdit();
			this.grvData = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.colDateTime = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colWorkerName = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colLot = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colHyp = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colGear = new DevExpress.XtraGrid.Columns.GridColumn();
			this.grdReport = new DevExpress.XtraGrid.GridControl();
			this.txtGear = new DevExpress.XtraEditors.TextEdit();
			this.txtHyp = new DevExpress.XtraEditors.TextEdit();
			this.colRiskDescription = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colRiskTime = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colDelayTime = new DevExpress.XtraGrid.Columns.GridColumn();
			((System.ComponentModel.ISupportInitialize)(this.txtLot.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemMemoEdit2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grvData)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grdReport)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGear.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtHyp.Properties)).BeginInit();
			this.SuspendLayout();
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(26, 17);
			this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(50, 20);
			this.label1.TabIndex = 0;
			this.label1.Text = "From:";
			// 
			// dtpFrom
			// 
			this.dtpFrom.CustomFormat = "dd/MM/yyyy";
			this.dtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
			this.dtpFrom.Location = new System.Drawing.Point(123, 12);
			this.dtpFrom.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
			this.dtpFrom.Name = "dtpFrom";
			this.dtpFrom.Size = new System.Drawing.Size(338, 26);
			this.dtpFrom.TabIndex = 0;
			this.dtpFrom.Tag = "";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(528, 17);
			this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(31, 20);
			this.label2.TabIndex = 2;
			this.label2.Text = "To:";
			// 
			// dtpTo
			// 
			this.dtpTo.CustomFormat = "dd/MM/yyyy";
			this.dtpTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
			this.dtpTo.Location = new System.Drawing.Point(614, 12);
			this.dtpTo.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
			this.dtpTo.Name = "dtpTo";
			this.dtpTo.Size = new System.Drawing.Size(354, 26);
			this.dtpTo.TabIndex = 1;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(26, 65);
			this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(84, 20);
			this.label3.TabIndex = 4;
			this.label3.Text = "MÃ GEAR";
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(26, 118);
			this.label4.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(70, 20);
			this.label4.TabIndex = 6;
			this.label4.Text = "MÃ HYP";
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(508, 65);
			this.label5.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(96, 20);
			this.label5.TabIndex = 8;
			this.label5.Text = "Lot gia công";
			// 
			// txtLot
			// 
			this.txtLot.Location = new System.Drawing.Point(614, 60);
			this.txtLot.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
			this.txtLot.Name = "txtLot";
			this.txtLot.Properties.Appearance.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.txtLot.Properties.Appearance.Options.UseFont = true;
			this.txtLot.Size = new System.Drawing.Size(356, 28);
			this.txtLot.TabIndex = 3;
			// 
			// btnWriteExcel
			// 
			this.btnWriteExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnWriteExcel.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnWriteExcel.Location = new System.Drawing.Point(1502, 12);
			this.btnWriteExcel.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
			this.btnWriteExcel.Name = "btnWriteExcel";
			this.btnWriteExcel.Size = new System.Drawing.Size(255, 132);
			this.btnWriteExcel.TabIndex = 6;
			this.btnWriteExcel.Text = "Ghi File Excel";
			this.btnWriteExcel.UseVisualStyleBackColor = true;
			this.btnWriteExcel.Click += new System.EventHandler(this.btnWriteExcel_Click);
			// 
			// btnGetReport
			// 
			this.btnGetReport.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnGetReport.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnGetReport.Location = new System.Drawing.Point(1256, 12);
			this.btnGetReport.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
			this.btnGetReport.Name = "btnGetReport";
			this.btnGetReport.Size = new System.Drawing.Size(237, 132);
			this.btnGetReport.TabIndex = 5;
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
            this.colWorkerName,
            this.colLot,
            this.colHyp,
            this.colGear});
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
			// colWorkerName
			// 
			this.colWorkerName.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.colWorkerName.AppearanceCell.Options.UseFont = true;
			this.colWorkerName.AppearanceCell.Options.UseTextOptions = true;
			this.colWorkerName.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colWorkerName.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colWorkerName.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colWorkerName.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.colWorkerName.AppearanceHeader.Options.UseFont = true;
			this.colWorkerName.AppearanceHeader.Options.UseTextOptions = true;
			this.colWorkerName.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colWorkerName.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colWorkerName.Caption = "Người Thao Tác";
			this.colWorkerName.FieldName = "WorkerCode";
			this.colWorkerName.Name = "colWorkerName";
			this.colWorkerName.OptionsColumn.AllowEdit = false;
			this.colWorkerName.OptionsColumn.AllowFocus = false;
			this.colWorkerName.OptionsColumn.AllowMerge = DevExpress.Utils.DefaultBoolean.False;
			this.colWorkerName.OptionsColumn.AllowMove = false;
			this.colWorkerName.Visible = true;
			this.colWorkerName.VisibleIndex = 1;
			this.colWorkerName.Width = 110;
			// 
			// colLot
			// 
			this.colLot.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.colLot.AppearanceCell.Options.UseFont = true;
			this.colLot.AppearanceCell.Options.UseTextOptions = true;
			this.colLot.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colLot.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colLot.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colLot.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.colLot.AppearanceHeader.Options.UseFont = true;
			this.colLot.Caption = "Lot Gia Công";
			this.colLot.ColumnEdit = this.repositoryItemMemoEdit2;
			this.colLot.FieldName = "LotGiaCong";
			this.colLot.Name = "colLot";
			this.colLot.OptionsColumn.AllowEdit = false;
			this.colLot.OptionsColumn.AllowFocus = false;
			this.colLot.OptionsColumn.AllowMerge = DevExpress.Utils.DefaultBoolean.False;
			this.colLot.OptionsColumn.AllowMove = false;
			this.colLot.Visible = true;
			this.colLot.VisibleIndex = 2;
			this.colLot.Width = 132;
			// 
			// colHyp
			// 
			this.colHyp.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.colHyp.AppearanceCell.Options.UseFont = true;
			this.colHyp.AppearanceCell.Options.UseTextOptions = true;
			this.colHyp.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colHyp.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colHyp.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colHyp.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.colHyp.AppearanceHeader.Options.UseFont = true;
			this.colHyp.AppearanceHeader.Options.UseTextOptions = true;
			this.colHyp.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colHyp.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colHyp.Caption = "Mã HYP";
			this.colHyp.FieldName = "Hyp";
			this.colHyp.Name = "colHyp";
			this.colHyp.OptionsColumn.AllowEdit = false;
			this.colHyp.OptionsColumn.AllowFocus = false;
			this.colHyp.OptionsColumn.AllowMerge = DevExpress.Utils.DefaultBoolean.False;
			this.colHyp.OptionsColumn.AllowMove = false;
			this.colHyp.Visible = true;
			this.colHyp.VisibleIndex = 3;
			this.colHyp.Width = 134;
			// 
			// colGear
			// 
			this.colGear.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.colGear.AppearanceCell.Options.UseFont = true;
			this.colGear.AppearanceCell.Options.UseTextOptions = true;
			this.colGear.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colGear.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colGear.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.colGear.AppearanceHeader.Options.UseFont = true;
			this.colGear.AppearanceHeader.Options.UseTextOptions = true;
			this.colGear.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colGear.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colGear.Caption = "Mã Gear";
			this.colGear.FieldName = "Gear";
			this.colGear.Name = "colGear";
			this.colGear.OptionsColumn.AllowEdit = false;
			this.colGear.OptionsColumn.AllowFocus = false;
			this.colGear.OptionsColumn.AllowMerge = DevExpress.Utils.DefaultBoolean.False;
			this.colGear.OptionsColumn.AllowMove = false;
			this.colGear.Visible = true;
			this.colGear.VisibleIndex = 4;
			this.colGear.Width = 108;
			// 
			// grdReport
			// 
			this.grdReport.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.grdReport.EmbeddedNavigator.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
			this.grdReport.Font = new System.Drawing.Font("Times New Roman", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.grdReport.Location = new System.Drawing.Point(3, 162);
			this.grdReport.MainView = this.grvData;
			this.grdReport.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
			this.grdReport.Name = "grdReport";
			this.grdReport.RepositoryItems.AddRange(new DevExpress.XtraEditors.Repository.RepositoryItem[] {
            this.repositoryItemMemoEdit2});
			this.grdReport.Size = new System.Drawing.Size(1760, 788);
			this.grdReport.TabIndex = 22;
			this.grdReport.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.grvData});
			// 
			// txtGear
			// 
			this.txtGear.Location = new System.Drawing.Point(123, 60);
			this.txtGear.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
			this.txtGear.Name = "txtGear";
			this.txtGear.Properties.Appearance.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.txtGear.Properties.Appearance.Options.UseFont = true;
			this.txtGear.Size = new System.Drawing.Size(340, 28);
			this.txtGear.TabIndex = 2;
			// 
			// txtHyp
			// 
			this.txtHyp.Location = new System.Drawing.Point(123, 108);
			this.txtHyp.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
			this.txtHyp.Name = "txtHyp";
			this.txtHyp.Properties.Appearance.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.txtHyp.Properties.Appearance.Options.UseFont = true;
			this.txtHyp.Size = new System.Drawing.Size(340, 28);
			this.txtHyp.TabIndex = 4;
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
			this.colRiskDescription.Width = 132;
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
			this.colDelayTime.Width = 90;
			// 
			// frmReportTester
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(1761, 954);
			this.Controls.Add(this.txtHyp);
			this.Controls.Add(this.txtGear);
			this.Controls.Add(this.btnGetReport);
			this.Controls.Add(this.btnWriteExcel);
			this.Controls.Add(this.grdReport);
			this.Controls.Add(this.txtLot);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.dtpTo);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.dtpFrom);
			this.Controls.Add(this.label1);
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.Name = "frmReportTester";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "BÁO CÁO DỮ LIỆU TESTER";
			this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
			this.Load += new System.EventHandler(this.frmExportAndon_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtLot.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemMemoEdit2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grvData)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grdReport)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGear.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtHyp.Properties)).EndInit();
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
        private DevExpress.XtraEditors.TextEdit txtLot;
        private System.Windows.Forms.Button btnWriteExcel;
        private System.Windows.Forms.Button btnGetReport;
        private System.Windows.Forms.SaveFileDialog sfdSaveDialog;
        private DevExpress.XtraEditors.Repository.RepositoryItemMemoEdit repositoryItemMemoEdit2;
        private DevExpress.XtraGrid.Views.Grid.GridView grvData;
        private DevExpress.XtraGrid.Columns.GridColumn colDateTime;
        private DevExpress.XtraGrid.Columns.GridColumn colWorkerName;
        private DevExpress.XtraGrid.Columns.GridColumn colLot;
        private DevExpress.XtraGrid.Columns.GridColumn colHyp;
        private DevExpress.XtraGrid.Columns.GridColumn colGear;
        private DevExpress.XtraGrid.GridControl grdReport;
        private DevExpress.XtraEditors.TextEdit txtGear;
        private DevExpress.XtraEditors.TextEdit txtHyp;
        private DevExpress.XtraGrid.Columns.GridColumn colRiskDescription;
        private DevExpress.XtraGrid.Columns.GridColumn colRiskTime;
        private DevExpress.XtraGrid.Columns.GridColumn colDelayTime;
    }
}