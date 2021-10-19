namespace BMS
{
    partial class frmKnifeWorkingH
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
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmKnifeWorkingH));
			this.grdData = new DevExpress.XtraGrid.GridControl();
			this.grvData = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.colStepID = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colStepCode = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colStepName = new DevExpress.XtraGrid.Columns.GridColumn();
			this.grdDetail = new DevExpress.XtraGrid.GridControl();
			this.grvDetail = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.colWorkingID = new DevExpress.XtraGrid.Columns.GridColumn();
			this.gridColumn2 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colWorkingName = new DevExpress.XtraGrid.Columns.GridColumn();
			this.repositoryItemMemoEdit1 = new DevExpress.XtraEditors.Repository.RepositoryItemMemoEdit();
			this.colValueTypeName = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colCheckValueType = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colValue = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colMin = new DevExpress.XtraGrid.Columns.GridColumn();
			this.repositoryItemTextEdit1 = new DevExpress.XtraEditors.Repository.RepositoryItemTextEdit();
			this.colMax = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colUnit = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colSTT = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colProductStepID = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colSortOrders = new DevExpress.XtraGrid.Columns.GridColumn();
			this.repositoryItemCheckEdit1 = new DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit();
			this.label1 = new System.Windows.Forms.Label();
			this.splitContainer1 = new System.Windows.Forms.SplitContainer();
			this.panel2 = new System.Windows.Forms.Panel();
			this.toolStrip1 = new System.Windows.Forms.ToolStrip();
			this.btnDeleteStep = new System.Windows.Forms.ToolStripButton();
			this.toolStripSeparator6 = new System.Windows.Forms.ToolStripSeparator();
			this.btnEditStep = new System.Windows.Forms.ToolStripButton();
			this.toolStripSeparator5 = new System.Windows.Forms.ToolStripSeparator();
			this.btnAddStep = new System.Windows.Forms.ToolStripButton();
			this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
			this.btnRestartCD = new System.Windows.Forms.ToolStripButton();
			this.panel4 = new System.Windows.Forms.Panel();
			this.chkAll = new System.Windows.Forms.CheckBox();
			this.mnuMenu = new System.Windows.Forms.ToolStrip();
			this.btnDelete = new System.Windows.Forms.ToolStripButton();
			this.toolStripSeparator4 = new System.Windows.Forms.ToolStripSeparator();
			this.btnEdit = new System.Windows.Forms.ToolStripButton();
			this.toolStripSeparator3 = new System.Windows.Forms.ToolStripSeparator();
			this.btnNew = new System.Windows.Forms.ToolStripButton();
			this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
			this.btnRestart = new System.Windows.Forms.ToolStripButton();
			this.btnSave = new System.Windows.Forms.ToolStripButton();
			this.panel3 = new System.Windows.Forms.Panel();
			this.colFrequency = new DevExpress.XtraGrid.Columns.GridColumn();
			((System.ComponentModel.ISupportInitialize)(this.grdData)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grvData)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grdDetail)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grvDetail)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemMemoEdit1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemTextEdit1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemCheckEdit1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
			this.splitContainer1.Panel1.SuspendLayout();
			this.splitContainer1.Panel2.SuspendLayout();
			this.splitContainer1.SuspendLayout();
			this.panel2.SuspendLayout();
			this.toolStrip1.SuspendLayout();
			this.panel4.SuspendLayout();
			this.mnuMenu.SuspendLayout();
			this.panel3.SuspendLayout();
			this.SuspendLayout();
			// 
			// grdData
			// 
			this.grdData.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.grdData.Location = new System.Drawing.Point(-1, 59);
			this.grdData.MainView = this.grvData;
			this.grdData.Name = "grdData";
			this.grdData.Size = new System.Drawing.Size(348, 618);
			this.grdData.TabIndex = 25;
			this.grdData.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.grvData});
			// 
			// grvData
			// 
			this.grvData.Appearance.FocusedRow.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(192)))), ((int)(((byte)(255)))));
			this.grvData.Appearance.FocusedRow.Options.UseBackColor = true;
			this.grvData.Appearance.HideSelectionRow.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(192)))), ((int)(((byte)(255)))));
			this.grvData.Appearance.HideSelectionRow.Options.UseBackColor = true;
			this.grvData.Appearance.SelectedRow.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(192)))), ((int)(((byte)(255)))));
			this.grvData.Appearance.SelectedRow.Options.UseBackColor = true;
			this.grvData.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
            this.colStepID,
            this.colStepCode,
            this.colStepName});
			this.grvData.GridControl = this.grdData;
			this.grvData.Name = "grvData";
			this.grvData.OptionsBehavior.Editable = false;
			this.grvData.OptionsView.ShowGroupPanel = false;
			this.grvData.FocusedRowChanged += new DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventHandler(this.grvData_FocusedRowChanged);
			this.grvData.DoubleClick += new System.EventHandler(this.grvData_DoubleClick);
			// 
			// colStepID
			// 
			this.colStepID.Caption = "ID";
			this.colStepID.FieldName = "ID";
			this.colStepID.Name = "colStepID";
			this.colStepID.OptionsColumn.AllowEdit = false;
			// 
			// colStepCode
			// 
			this.colStepCode.AppearanceCell.Options.UseTextOptions = true;
			this.colStepCode.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colStepCode.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colStepCode.AppearanceHeader.Options.UseFont = true;
			this.colStepCode.AppearanceHeader.Options.UseTextOptions = true;
			this.colStepCode.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colStepCode.Caption = "Công đoạn";
			this.colStepCode.FieldName = "ProductStepCode";
			this.colStepCode.Name = "colStepCode";
			this.colStepCode.OptionsColumn.AllowEdit = false;
			this.colStepCode.OptionsColumn.AllowFocus = false;
			this.colStepCode.Visible = true;
			this.colStepCode.VisibleIndex = 0;
			this.colStepCode.Width = 72;
			// 
			// colStepName
			// 
			this.colStepName.AppearanceCell.Options.UseTextOptions = true;
			this.colStepName.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colStepName.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colStepName.AppearanceHeader.Options.UseFont = true;
			this.colStepName.AppearanceHeader.Options.UseTextOptions = true;
			this.colStepName.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colStepName.Caption = "Tên công đoạn";
			this.colStepName.FieldName = "Description";
			this.colStepName.Name = "colStepName";
			this.colStepName.OptionsColumn.AllowEdit = false;
			this.colStepName.OptionsColumn.AllowFocus = false;
			this.colStepName.Visible = true;
			this.colStepName.VisibleIndex = 1;
			this.colStepName.Width = 250;
			// 
			// grdDetail
			// 
			this.grdDetail.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.grdDetail.Location = new System.Drawing.Point(0, 59);
			this.grdDetail.MainView = this.grvDetail;
			this.grdDetail.Name = "grdDetail";
			this.grdDetail.RepositoryItems.AddRange(new DevExpress.XtraEditors.Repository.RepositoryItem[] {
            this.repositoryItemMemoEdit1,
            this.repositoryItemTextEdit1,
            this.repositoryItemCheckEdit1});
			this.grdDetail.Size = new System.Drawing.Size(813, 622);
			this.grdDetail.TabIndex = 25;
			this.grdDetail.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.grvDetail});
			this.grdDetail.Click += new System.EventHandler(this.grdDetail_Click);
			// 
			// grvDetail
			// 
			this.grvDetail.Appearance.FocusedRow.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(192)))), ((int)(((byte)(255)))));
			this.grvDetail.Appearance.FocusedRow.BackColor2 = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(192)))), ((int)(((byte)(255)))));
			this.grvDetail.Appearance.FocusedRow.Options.UseBackColor = true;
			this.grvDetail.Appearance.HideSelectionRow.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(192)))), ((int)(((byte)(255)))));
			this.grvDetail.Appearance.HideSelectionRow.BackColor2 = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(192)))), ((int)(((byte)(255)))));
			this.grvDetail.Appearance.HideSelectionRow.Options.UseBackColor = true;
			this.grvDetail.Appearance.SelectedRow.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(192)))), ((int)(((byte)(255)))));
			this.grvDetail.Appearance.SelectedRow.BackColor2 = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(192)))), ((int)(((byte)(255)))));
			this.grvDetail.Appearance.SelectedRow.Options.UseBackColor = true;
			this.grvDetail.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
            this.colWorkingID,
            this.gridColumn2,
            this.colWorkingName,
            this.colValueTypeName,
            this.colCheckValueType,
            this.colValue,
            this.colMin,
            this.colMax,
            this.colUnit,
            this.colSTT,
            this.colProductStepID,
            this.colSortOrders,
            this.colFrequency});
			this.grvDetail.GridControl = this.grdDetail;
			this.grvDetail.GroupCount = 1;
			this.grvDetail.Name = "grvDetail";
			this.grvDetail.OptionsBehavior.AutoExpandAllGroups = true;
			this.grvDetail.OptionsBehavior.Editable = false;
			this.grvDetail.OptionsCustomization.AllowRowSizing = true;
			this.grvDetail.OptionsFind.ShowCloseButton = false;
			this.grvDetail.OptionsSelection.MultiSelect = true;
			this.grvDetail.OptionsView.ColumnAutoWidth = false;
			this.grvDetail.OptionsView.RowAutoHeight = true;
			this.grvDetail.OptionsView.ShowAutoFilterRow = true;
			this.grvDetail.OptionsView.ShowFooter = true;
			this.grvDetail.OptionsView.ShowGroupPanel = false;
			this.grvDetail.SortInfo.AddRange(new DevExpress.XtraGrid.Columns.GridColumnSortInfo[] {
            new DevExpress.XtraGrid.Columns.GridColumnSortInfo(this.gridColumn2, DevExpress.Data.ColumnSortOrder.Ascending)});
			this.grvDetail.CustomDrawCell += new DevExpress.XtraGrid.Views.Base.RowCellCustomDrawEventHandler(this.grvDetail_CustomDrawCell);
			this.grvDetail.DoubleClick += new System.EventHandler(this.grdDetail_DoubleClick);
			// 
			// colWorkingID
			// 
			this.colWorkingID.Caption = "ID";
			this.colWorkingID.FieldName = "ID";
			this.colWorkingID.Name = "colWorkingID";
			this.colWorkingID.OptionsColumn.AllowEdit = false;
			// 
			// gridColumn2
			// 
			this.gridColumn2.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.gridColumn2.AppearanceHeader.Options.UseFont = true;
			this.gridColumn2.AppearanceHeader.Options.UseTextOptions = true;
			this.gridColumn2.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.gridColumn2.Caption = "Công đoạn";
			this.gridColumn2.FieldName = "ProductStepCodes";
			this.gridColumn2.Name = "gridColumn2";
			this.gridColumn2.OptionsColumn.AllowEdit = false;
			this.gridColumn2.OptionsColumn.AllowFocus = false;
			this.gridColumn2.Width = 247;
			// 
			// colWorkingName
			// 
			this.colWorkingName.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colWorkingName.AppearanceHeader.Options.UseFont = true;
			this.colWorkingName.AppearanceHeader.Options.UseTextOptions = true;
			this.colWorkingName.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colWorkingName.Caption = "Mục kiểm tra";
			this.colWorkingName.ColumnEdit = this.repositoryItemMemoEdit1;
			this.colWorkingName.FieldName = "WorkingName";
			this.colWorkingName.Name = "colWorkingName";
			this.colWorkingName.OptionsColumn.AllowSize = false;
			this.colWorkingName.Visible = true;
			this.colWorkingName.VisibleIndex = 1;
			this.colWorkingName.Width = 303;
			// 
			// repositoryItemMemoEdit1
			// 
			this.repositoryItemMemoEdit1.Appearance.Options.UseTextOptions = true;
			this.repositoryItemMemoEdit1.Appearance.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.repositoryItemMemoEdit1.Name = "repositoryItemMemoEdit1";
			// 
			// colValueTypeName
			// 
			this.colValueTypeName.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colValueTypeName.AppearanceHeader.Options.UseFont = true;
			this.colValueTypeName.AppearanceHeader.Options.UseTextOptions = true;
			this.colValueTypeName.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colValueTypeName.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colValueTypeName.Caption = "Ghi chép";
			this.colValueTypeName.FieldName = "ValueTypeName";
			this.colValueTypeName.Name = "colValueTypeName";
			this.colValueTypeName.OptionsColumn.AllowEdit = false;
			this.colValueTypeName.Visible = true;
			this.colValueTypeName.VisibleIndex = 2;
			this.colValueTypeName.Width = 70;
			// 
			// colCheckValueType
			// 
			this.colCheckValueType.AppearanceCell.Options.UseTextOptions = true;
			this.colCheckValueType.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colCheckValueType.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colCheckValueType.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colCheckValueType.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.colCheckValueType.AppearanceHeader.Options.UseFont = true;
			this.colCheckValueType.AppearanceHeader.Options.UseTextOptions = true;
			this.colCheckValueType.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colCheckValueType.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colCheckValueType.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colCheckValueType.Caption = "Loại giá trị";
			this.colCheckValueType.FieldName = "CheckValueTypes";
			this.colCheckValueType.Name = "colCheckValueType";
			this.colCheckValueType.Visible = true;
			this.colCheckValueType.VisibleIndex = 3;
			// 
			// colValue
			// 
			this.colValue.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colValue.AppearanceHeader.Options.UseFont = true;
			this.colValue.AppearanceHeader.Options.UseTextOptions = true;
			this.colValue.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colValue.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colValue.Caption = "Giá trị tiêu chuẩn";
			this.colValue.FieldName = "PeriodValue";
			this.colValue.Name = "colValue";
			this.colValue.Visible = true;
			this.colValue.VisibleIndex = 5;
			this.colValue.Width = 114;
			// 
			// colMin
			// 
			this.colMin.AppearanceCell.Options.UseTextOptions = true;
			this.colMin.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colMin.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colMin.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colMin.AppearanceHeader.Options.UseFont = true;
			this.colMin.AppearanceHeader.Options.UseTextOptions = true;
			this.colMin.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colMin.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colMin.Caption = "Min";
			this.colMin.ColumnEdit = this.repositoryItemTextEdit1;
			this.colMin.DisplayFormat.FormatString = "n3";
			this.colMin.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.colMin.FieldName = "MinValue";
			this.colMin.Name = "colMin";
			this.colMin.Visible = true;
			this.colMin.VisibleIndex = 6;
			// 
			// repositoryItemTextEdit1
			// 
			this.repositoryItemTextEdit1.AutoHeight = false;
			this.repositoryItemTextEdit1.DisplayFormat.FormatString = "n3";
			this.repositoryItemTextEdit1.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.repositoryItemTextEdit1.EditFormat.FormatString = "n3";
			this.repositoryItemTextEdit1.EditFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.repositoryItemTextEdit1.Mask.EditMask = "n3";
			this.repositoryItemTextEdit1.Mask.MaskType = DevExpress.XtraEditors.Mask.MaskType.Numeric;
			this.repositoryItemTextEdit1.Name = "repositoryItemTextEdit1";
			// 
			// colMax
			// 
			this.colMax.AppearanceCell.Options.UseTextOptions = true;
			this.colMax.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colMax.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colMax.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colMax.AppearanceHeader.Options.UseFont = true;
			this.colMax.AppearanceHeader.Options.UseTextOptions = true;
			this.colMax.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colMax.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colMax.Caption = "Max";
			this.colMax.ColumnEdit = this.repositoryItemTextEdit1;
			this.colMax.DisplayFormat.FormatString = "n3";
			this.colMax.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.colMax.FieldName = "MaxValue";
			this.colMax.Name = "colMax";
			this.colMax.Visible = true;
			this.colMax.VisibleIndex = 7;
			this.colMax.Width = 83;
			// 
			// colUnit
			// 
			this.colUnit.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colUnit.AppearanceHeader.Options.UseFont = true;
			this.colUnit.AppearanceHeader.Options.UseTextOptions = true;
			this.colUnit.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colUnit.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colUnit.Caption = "Đơn vị tính";
			this.colUnit.FieldName = "Unit";
			this.colUnit.Name = "colUnit";
			this.colUnit.OptionsColumn.AllowEdit = false;
			this.colUnit.Visible = true;
			this.colUnit.VisibleIndex = 4;
			this.colUnit.Width = 72;
			// 
			// colSTT
			// 
			this.colSTT.AppearanceCell.Options.UseTextOptions = true;
			this.colSTT.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colSTT.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colSTT.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colSTT.AppearanceHeader.Options.UseFont = true;
			this.colSTT.AppearanceHeader.Options.UseTextOptions = true;
			this.colSTT.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colSTT.Caption = "STT";
			this.colSTT.FieldName = "SortOrder";
			this.colSTT.Name = "colSTT";
			this.colSTT.OptionsColumn.AllowEdit = false;
			this.colSTT.Visible = true;
			this.colSTT.VisibleIndex = 0;
			this.colSTT.Width = 57;
			// 
			// colProductStepID
			// 
			this.colProductStepID.Caption = "ProductStepID";
			this.colProductStepID.FieldName = "ProductStepID";
			this.colProductStepID.Name = "colProductStepID";
			this.colProductStepID.OptionsColumn.AllowEdit = false;
			// 
			// colSortOrders
			// 
			this.colSortOrders.Caption = "SortOrders";
			this.colSortOrders.FieldName = "SortOrders";
			this.colSortOrders.Name = "colSortOrders";
			// 
			// repositoryItemCheckEdit1
			// 
			this.repositoryItemCheckEdit1.AutoHeight = false;
			this.repositoryItemCheckEdit1.Caption = "Check";
			this.repositoryItemCheckEdit1.Name = "repositoryItemCheckEdit1";
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.label1.Location = new System.Drawing.Point(7, 33);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(169, 13);
			this.label1.TabIndex = 28;
			this.label1.Text = "Danh sách mục cần kiểm tra";
			// 
			// splitContainer1
			// 
			this.splitContainer1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.splitContainer1.Location = new System.Drawing.Point(0, 0);
			this.splitContainer1.Name = "splitContainer1";
			// 
			// splitContainer1.Panel1
			// 
			this.splitContainer1.Panel1.Controls.Add(this.panel2);
			this.splitContainer1.Panel1.Controls.Add(this.grdData);
			// 
			// splitContainer1.Panel2
			// 
			this.splitContainer1.Panel2.Controls.Add(this.panel4);
			this.splitContainer1.Panel2.Controls.Add(this.panel3);
			this.splitContainer1.Size = new System.Drawing.Size(1167, 683);
			this.splitContainer1.SplitterDistance = 349;
			this.splitContainer1.TabIndex = 29;
			// 
			// panel2
			// 
			this.panel2.Controls.Add(this.toolStrip1);
			this.panel2.Dock = System.Windows.Forms.DockStyle.Top;
			this.panel2.Location = new System.Drawing.Point(0, 0);
			this.panel2.Margin = new System.Windows.Forms.Padding(2);
			this.panel2.Name = "panel2";
			this.panel2.Size = new System.Drawing.Size(347, 58);
			this.panel2.TabIndex = 31;
			// 
			// toolStrip1
			// 
			this.toolStrip1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.toolStrip1.AutoSize = false;
			this.toolStrip1.BackColor = System.Drawing.Color.WhiteSmoke;
			this.toolStrip1.Dock = System.Windows.Forms.DockStyle.None;
			this.toolStrip1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.toolStrip1.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden;
			this.toolStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
			this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.btnDeleteStep,
            this.toolStripSeparator6,
            this.btnEditStep,
            this.toolStripSeparator5,
            this.btnAddStep,
            this.toolStripSeparator2,
            this.btnRestartCD});
			this.toolStrip1.LayoutStyle = System.Windows.Forms.ToolStripLayoutStyle.Flow;
			this.toolStrip1.Location = new System.Drawing.Point(112, 10);
			this.toolStrip1.Name = "toolStrip1";
			this.toolStrip1.RenderMode = System.Windows.Forms.ToolStripRenderMode.Professional;
			this.toolStrip1.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
			this.toolStrip1.Size = new System.Drawing.Size(235, 46);
			this.toolStrip1.TabIndex = 30;
			this.toolStrip1.Text = "toolStrip2";
			// 
			// btnDeleteStep
			// 
			this.btnDeleteStep.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnDeleteStep.Image = ((System.Drawing.Image)(resources.GetObject("btnDeleteStep.Image")));
			this.btnDeleteStep.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnDeleteStep.Name = "btnDeleteStep";
			this.btnDeleteStep.Size = new System.Drawing.Size(32, 37);
			this.btnDeleteStep.Tag = "frmProductWorkingH_DeleteCD";
			this.btnDeleteStep.Text = "Xóa";
			this.btnDeleteStep.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnDeleteStep.Click += new System.EventHandler(this.btnDeleteStep_Click);
			// 
			// toolStripSeparator6
			// 
			this.toolStripSeparator6.Name = "toolStripSeparator6";
			this.toolStripSeparator6.Size = new System.Drawing.Size(6, 23);
			// 
			// btnEditStep
			// 
			this.btnEditStep.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnEditStep.Image = ((System.Drawing.Image)(resources.GetObject("btnEditStep.Image")));
			this.btnEditStep.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnEditStep.Name = "btnEditStep";
			this.btnEditStep.Size = new System.Drawing.Size(33, 37);
			this.btnEditStep.Tag = "frmProductWorkingH_EditCD";
			this.btnEditStep.Text = "Sửa";
			this.btnEditStep.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnEditStep.Click += new System.EventHandler(this.btnEditStep_Click);
			// 
			// toolStripSeparator5
			// 
			this.toolStripSeparator5.Name = "toolStripSeparator5";
			this.toolStripSeparator5.Size = new System.Drawing.Size(6, 23);
			// 
			// btnAddStep
			// 
			this.btnAddStep.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnAddStep.Image = ((System.Drawing.Image)(resources.GetObject("btnAddStep.Image")));
			this.btnAddStep.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnAddStep.Name = "btnAddStep";
			this.btnAddStep.Size = new System.Drawing.Size(43, 37);
			this.btnAddStep.Tag = "frmProductWorkingH_AddCD";
			this.btnAddStep.Text = "Thêm";
			this.btnAddStep.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnAddStep.Click += new System.EventHandler(this.btnAddStep_Click);
			// 
			// toolStripSeparator2
			// 
			this.toolStripSeparator2.Name = "toolStripSeparator2";
			this.toolStripSeparator2.Size = new System.Drawing.Size(6, 23);
			// 
			// btnRestartCD
			// 
			this.btnRestartCD.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnRestartCD.Image = ((System.Drawing.Image)(resources.GetObject("btnRestartCD.Image")));
			this.btnRestartCD.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnRestartCD.Name = "btnRestartCD";
			this.btnRestartCD.Size = new System.Drawing.Size(59, 37);
			this.btnRestartCD.Tag = "";
			this.btnRestartCD.Text = "Làm mới";
			this.btnRestartCD.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnRestartCD.Click += new System.EventHandler(this.btnRestartCD_Click);
			// 
			// panel4
			// 
			this.panel4.Controls.Add(this.chkAll);
			this.panel4.Controls.Add(this.mnuMenu);
			this.panel4.Controls.Add(this.label1);
			this.panel4.Dock = System.Windows.Forms.DockStyle.Top;
			this.panel4.Location = new System.Drawing.Point(0, 0);
			this.panel4.Margin = new System.Windows.Forms.Padding(2);
			this.panel4.Name = "panel4";
			this.panel4.Size = new System.Drawing.Size(812, 56);
			this.panel4.TabIndex = 33;
			// 
			// chkAll
			// 
			this.chkAll.AutoSize = true;
			this.chkAll.Location = new System.Drawing.Point(10, 10);
			this.chkAll.Name = "chkAll";
			this.chkAll.Size = new System.Drawing.Size(92, 17);
			this.chkAll.TabIndex = 30;
			this.chkAll.Text = "Hiển thị tất cả";
			this.chkAll.UseVisualStyleBackColor = true;
			this.chkAll.CheckedChanged += new System.EventHandler(this.chkAll_CheckedChanged);
			// 
			// mnuMenu
			// 
			this.mnuMenu.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.mnuMenu.AutoSize = false;
			this.mnuMenu.BackColor = System.Drawing.Color.WhiteSmoke;
			this.mnuMenu.Dock = System.Windows.Forms.DockStyle.None;
			this.mnuMenu.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.mnuMenu.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden;
			this.mnuMenu.ImageScalingSize = new System.Drawing.Size(20, 20);
			this.mnuMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.btnDelete,
            this.toolStripSeparator4,
            this.btnEdit,
            this.toolStripSeparator3,
            this.btnNew,
            this.toolStripSeparator1,
            this.btnRestart,
            this.btnSave});
			this.mnuMenu.LayoutStyle = System.Windows.Forms.ToolStripLayoutStyle.Flow;
			this.mnuMenu.Location = new System.Drawing.Point(552, 10);
			this.mnuMenu.Name = "mnuMenu";
			this.mnuMenu.RenderMode = System.Windows.Forms.ToolStripRenderMode.Professional;
			this.mnuMenu.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
			this.mnuMenu.Size = new System.Drawing.Size(258, 37);
			this.mnuMenu.TabIndex = 29;
			this.mnuMenu.Text = "toolStrip2";
			// 
			// btnDelete
			// 
			this.btnDelete.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnDelete.Image = ((System.Drawing.Image)(resources.GetObject("btnDelete.Image")));
			this.btnDelete.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnDelete.Name = "btnDelete";
			this.btnDelete.Size = new System.Drawing.Size(32, 37);
			this.btnDelete.Tag = "frmProductWorkingH_DeleteDS";
			this.btnDelete.Text = "Xóa";
			this.btnDelete.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnDelete.Click += new System.EventHandler(this.btnDelete_Click);
			// 
			// toolStripSeparator4
			// 
			this.toolStripSeparator4.Name = "toolStripSeparator4";
			this.toolStripSeparator4.Size = new System.Drawing.Size(6, 23);
			// 
			// btnEdit
			// 
			this.btnEdit.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnEdit.Image = ((System.Drawing.Image)(resources.GetObject("btnEdit.Image")));
			this.btnEdit.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnEdit.Name = "btnEdit";
			this.btnEdit.Size = new System.Drawing.Size(33, 37);
			this.btnEdit.Tag = "frmProductWorkingH_EditDS";
			this.btnEdit.Text = "Sửa";
			this.btnEdit.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnEdit.Click += new System.EventHandler(this.btnEdit_Click);
			// 
			// toolStripSeparator3
			// 
			this.toolStripSeparator3.Name = "toolStripSeparator3";
			this.toolStripSeparator3.Size = new System.Drawing.Size(6, 23);
			// 
			// btnNew
			// 
			this.btnNew.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnNew.Image = ((System.Drawing.Image)(resources.GetObject("btnNew.Image")));
			this.btnNew.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnNew.Name = "btnNew";
			this.btnNew.Size = new System.Drawing.Size(43, 37);
			this.btnNew.Tag = "frmProductWorkingH_AddDS";
			this.btnNew.Text = "Thêm";
			this.btnNew.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnNew.Click += new System.EventHandler(this.btnNew_Click);
			// 
			// toolStripSeparator1
			// 
			this.toolStripSeparator1.Name = "toolStripSeparator1";
			this.toolStripSeparator1.Size = new System.Drawing.Size(6, 23);
			// 
			// btnRestart
			// 
			this.btnRestart.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnRestart.Image = ((System.Drawing.Image)(resources.GetObject("btnRestart.Image")));
			this.btnRestart.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnRestart.Name = "btnRestart";
			this.btnRestart.Size = new System.Drawing.Size(59, 37);
			this.btnRestart.Tag = "";
			this.btnRestart.Text = "Làm mới";
			this.btnRestart.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnRestart.Click += new System.EventHandler(this.btnRestart_Click);
			// 
			// btnSave
			// 
			this.btnSave.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnSave.Image = ((System.Drawing.Image)(resources.GetObject("btnSave.Image")));
			this.btnSave.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnSave.Name = "btnSave";
			this.btnSave.Size = new System.Drawing.Size(30, 37);
			this.btnSave.Text = "Cất";
			this.btnSave.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnSave.Visible = false;
			this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
			// 
			// panel3
			// 
			this.panel3.Controls.Add(this.grdDetail);
			this.panel3.Dock = System.Windows.Forms.DockStyle.Fill;
			this.panel3.Location = new System.Drawing.Point(0, 0);
			this.panel3.Margin = new System.Windows.Forms.Padding(2);
			this.panel3.Name = "panel3";
			this.panel3.Size = new System.Drawing.Size(812, 681);
			this.panel3.TabIndex = 32;
			// 
			// colFrequency
			// 
			this.colFrequency.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colFrequency.AppearanceHeader.Options.UseFont = true;
			this.colFrequency.AppearanceHeader.Options.UseTextOptions = true;
			this.colFrequency.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colFrequency.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colFrequency.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colFrequency.Caption = "Tần suất";
			this.colFrequency.FieldName = "Frequency";
			this.colFrequency.Name = "colFrequency";
			// 
			// frmKnifeWorkingH
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(1167, 683);
			this.Controls.Add(this.splitContainer1);
			this.Margin = new System.Windows.Forms.Padding(4);
			this.Name = "frmKnifeWorkingH";
			this.Text = "CHI TIẾT MỤC KIỂM TRA DAO && JIG && KHUÔN";
			this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
			this.Load += new System.EventHandler(this.frmProductWorking_Load);
			((System.ComponentModel.ISupportInitialize)(this.grdData)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grvData)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grdDetail)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grvDetail)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemMemoEdit1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemTextEdit1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemCheckEdit1)).EndInit();
			this.splitContainer1.Panel1.ResumeLayout(false);
			this.splitContainer1.Panel2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
			this.splitContainer1.ResumeLayout(false);
			this.panel2.ResumeLayout(false);
			this.toolStrip1.ResumeLayout(false);
			this.toolStrip1.PerformLayout();
			this.panel4.ResumeLayout(false);
			this.panel4.PerformLayout();
			this.mnuMenu.ResumeLayout(false);
			this.mnuMenu.PerformLayout();
			this.panel3.ResumeLayout(false);
			this.ResumeLayout(false);

        }

        #endregion
        private DevExpress.XtraGrid.GridControl grdData;
        private DevExpress.XtraGrid.Views.Grid.GridView grvData;
        private DevExpress.XtraGrid.Columns.GridColumn colStepID;
        private DevExpress.XtraGrid.Columns.GridColumn colStepCode;
        private DevExpress.XtraGrid.Columns.GridColumn colStepName;
        private DevExpress.XtraGrid.GridControl grdDetail;
        private DevExpress.XtraGrid.Views.Grid.GridView grvDetail;
        private DevExpress.XtraGrid.Columns.GridColumn colWorkingID;
        private DevExpress.XtraGrid.Columns.GridColumn gridColumn2;
        private DevExpress.XtraGrid.Columns.GridColumn colWorkingName;
        private System.Windows.Forms.Label label1;
        private DevExpress.XtraGrid.Columns.GridColumn colValueTypeName;
        private DevExpress.XtraGrid.Columns.GridColumn colValue;
        private DevExpress.XtraGrid.Columns.GridColumn colMin;
        private DevExpress.XtraGrid.Columns.GridColumn colMax;
        private DevExpress.XtraGrid.Columns.GridColumn colUnit;
        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.ToolStrip mnuMenu;
        private System.Windows.Forms.ToolStripButton btnNew;
        private System.Windows.Forms.ToolStripButton btnEdit;
        private System.Windows.Forms.ToolStripButton btnDelete;
        private DevExpress.XtraEditors.Repository.RepositoryItemMemoEdit repositoryItemMemoEdit1;
        private System.Windows.Forms.ToolStrip toolStrip1;
        private System.Windows.Forms.ToolStripButton btnAddStep;
        private System.Windows.Forms.ToolStripButton btnEditStep;
        private System.Windows.Forms.ToolStripButton btnDeleteStep;
        private DevExpress.XtraGrid.Columns.GridColumn colSTT;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripButton btnSave;
        private DevExpress.XtraEditors.Repository.RepositoryItemTextEdit repositoryItemTextEdit1;
        private DevExpress.XtraGrid.Columns.GridColumn colProductStepID;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Panel panel4;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.ToolStripButton btnRestart;
        private System.Windows.Forms.ToolStripButton btnRestartCD;
        private DevExpress.XtraGrid.Columns.GridColumn colSortOrders;
        private System.Windows.Forms.CheckBox chkAll;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator6;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator5;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator4;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator3;
        private DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit repositoryItemCheckEdit1;
        private DevExpress.XtraGrid.Columns.GridColumn colCheckValueType;
		private DevExpress.XtraGrid.Columns.GridColumn colFrequency;
	}
}