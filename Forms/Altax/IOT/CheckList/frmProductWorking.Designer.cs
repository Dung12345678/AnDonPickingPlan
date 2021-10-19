namespace BMS
{
    partial class frmProductWorking
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
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmProductWorking));
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
			this.gridColumn4 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colValue = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colMin = new DevExpress.XtraGrid.Columns.GridColumn();
			this.repositoryItemTextEdit1 = new DevExpress.XtraEditors.Repository.RepositoryItemTextEdit();
			this.colMax = new DevExpress.XtraGrid.Columns.GridColumn();
			this.gridColumn8 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colSTT = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colProductStepID = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colIsHidden = new DevExpress.XtraGrid.Columns.GridColumn();
			this.repositoryItemCheckEdit1 = new DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit();
			this.colCheckValueTypes = new DevExpress.XtraGrid.Columns.GridColumn();
			this.label1 = new System.Windows.Forms.Label();
			this.splitContainer1 = new System.Windows.Forms.SplitContainer();
			this.toolStrip1 = new System.Windows.Forms.ToolStrip();
			this.btnDeleteStep = new System.Windows.Forms.ToolStripButton();
			this.btnEditStep = new System.Windows.Forms.ToolStripButton();
			this.btnAddStep = new System.Windows.Forms.ToolStripButton();
			this.btnRestartCD = new System.Windows.Forms.ToolStripButton();
			this.chkAll = new System.Windows.Forms.CheckBox();
			this.mnuMenu = new System.Windows.Forms.ToolStrip();
			this.btnDelete = new System.Windows.Forms.ToolStripButton();
			this.btnEdit = new System.Windows.Forms.ToolStripButton();
			this.btnNew = new System.Windows.Forms.ToolStripButton();
			this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
			this.btnSave = new System.Windows.Forms.ToolStripButton();
			this.toolStripButton1 = new System.Windows.Forms.ToolStripButton();
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
			this.toolStrip1.SuspendLayout();
			this.mnuMenu.SuspendLayout();
			this.SuspendLayout();
			// 
			// grdData
			// 
			this.grdData.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.grdData.Location = new System.Drawing.Point(5, 51);
			this.grdData.MainView = this.grvData;
			this.grdData.Name = "grdData";
			this.grdData.Size = new System.Drawing.Size(340, 627);
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
			this.colStepCode.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colStepCode.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colStepCode.AppearanceHeader.Options.UseBackColor = true;
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
			this.colStepName.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colStepName.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colStepName.AppearanceHeader.Options.UseBackColor = true;
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
			this.grdDetail.Location = new System.Drawing.Point(3, 51);
			this.grdDetail.MainView = this.grvDetail;
			this.grdDetail.Name = "grdDetail";
			this.grdDetail.RepositoryItems.AddRange(new DevExpress.XtraEditors.Repository.RepositoryItem[] {
            this.repositoryItemMemoEdit1,
            this.repositoryItemTextEdit1,
            this.repositoryItemCheckEdit1});
			this.grdDetail.Size = new System.Drawing.Size(805, 627);
			this.grdDetail.TabIndex = 25;
			this.grdDetail.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.grvDetail});
			// 
			// grvDetail
			// 
			this.grvDetail.Appearance.FocusedRow.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(192)))), ((int)(((byte)(255)))));
			this.grvDetail.Appearance.FocusedRow.Options.UseBackColor = true;
			this.grvDetail.Appearance.HideSelectionRow.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(192)))), ((int)(((byte)(255)))));
			this.grvDetail.Appearance.HideSelectionRow.Options.UseBackColor = true;
			this.grvDetail.Appearance.SelectedRow.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(192)))), ((int)(((byte)(255)))));
			this.grvDetail.Appearance.SelectedRow.Options.UseBackColor = true;
			this.grvDetail.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
            this.colWorkingID,
            this.gridColumn2,
            this.colWorkingName,
            this.gridColumn4,
            this.colValue,
            this.colMin,
            this.colMax,
            this.gridColumn8,
            this.colSTT,
            this.colProductStepID,
            this.colIsHidden,
            this.colCheckValueTypes});
			this.grvDetail.GridControl = this.grdDetail;
			this.grvDetail.GroupCount = 1;
			this.grvDetail.Name = "grvDetail";
			this.grvDetail.OptionsBehavior.AutoExpandAllGroups = true;
			this.grvDetail.OptionsCustomization.AllowRowSizing = true;
			this.grvDetail.OptionsView.ColumnAutoWidth = false;
			this.grvDetail.OptionsView.RowAutoHeight = true;
			this.grvDetail.OptionsView.ShowGroupPanel = false;
			this.grvDetail.SortInfo.AddRange(new DevExpress.XtraGrid.Columns.GridColumnSortInfo[] {
            new DevExpress.XtraGrid.Columns.GridColumnSortInfo(this.gridColumn2, DevExpress.Data.ColumnSortOrder.Ascending)});
			this.grvDetail.DoubleClick += new System.EventHandler(this.grdDetail_DoubleClick);
			// 
			// colWorkingID
			// 
			this.colWorkingID.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colWorkingID.AppearanceHeader.Options.UseBackColor = true;
			this.colWorkingID.Caption = "ID";
			this.colWorkingID.FieldName = "ID";
			this.colWorkingID.Name = "colWorkingID";
			this.colWorkingID.OptionsColumn.AllowEdit = false;
			// 
			// gridColumn2
			// 
			this.gridColumn2.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.gridColumn2.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.gridColumn2.AppearanceHeader.Options.UseBackColor = true;
			this.gridColumn2.AppearanceHeader.Options.UseFont = true;
			this.gridColumn2.AppearanceHeader.Options.UseTextOptions = true;
			this.gridColumn2.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.gridColumn2.Caption = "Công đoạn";
			this.gridColumn2.FieldName = "ProductStepCode";
			this.gridColumn2.Name = "gridColumn2";
			this.gridColumn2.OptionsColumn.AllowEdit = false;
			this.gridColumn2.OptionsColumn.AllowFocus = false;
			this.gridColumn2.Width = 247;
			// 
			// colWorkingName
			// 
			this.colWorkingName.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colWorkingName.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colWorkingName.AppearanceHeader.Options.UseBackColor = true;
			this.colWorkingName.AppearanceHeader.Options.UseFont = true;
			this.colWorkingName.AppearanceHeader.Options.UseTextOptions = true;
			this.colWorkingName.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colWorkingName.Caption = "Mục kiểm tra";
			this.colWorkingName.ColumnEdit = this.repositoryItemMemoEdit1;
			this.colWorkingName.FieldName = "WorkingName";
			this.colWorkingName.Name = "colWorkingName";
			this.colWorkingName.OptionsColumn.AllowEdit = false;
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
			// gridColumn4
			// 
			this.gridColumn4.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.gridColumn4.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.gridColumn4.AppearanceHeader.Options.UseBackColor = true;
			this.gridColumn4.AppearanceHeader.Options.UseFont = true;
			this.gridColumn4.AppearanceHeader.Options.UseTextOptions = true;
			this.gridColumn4.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.gridColumn4.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.gridColumn4.Caption = "Ghi chép";
			this.gridColumn4.FieldName = "ValueTypeName";
			this.gridColumn4.Name = "gridColumn4";
			this.gridColumn4.OptionsColumn.AllowEdit = false;
			this.gridColumn4.Visible = true;
			this.gridColumn4.VisibleIndex = 3;
			this.gridColumn4.Width = 70;
			// 
			// colValue
			// 
			this.colValue.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colValue.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colValue.AppearanceHeader.Options.UseBackColor = true;
			this.colValue.AppearanceHeader.Options.UseFont = true;
			this.colValue.AppearanceHeader.Options.UseTextOptions = true;
			this.colValue.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colValue.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colValue.Caption = "Giá trị tiêu chuẩn";
			this.colValue.FieldName = "PeriodValue";
			this.colValue.Name = "colValue";
			this.colValue.OptionsColumn.AllowEdit = false;
			this.colValue.Visible = true;
			this.colValue.VisibleIndex = 6;
			this.colValue.Width = 114;
			// 
			// colMin
			// 
			this.colMin.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colMin.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colMin.AppearanceHeader.Options.UseBackColor = true;
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
			this.colMin.OptionsColumn.AllowEdit = false;
			this.colMin.Visible = true;
			this.colMin.VisibleIndex = 7;
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
			this.colMax.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colMax.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colMax.AppearanceHeader.Options.UseBackColor = true;
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
			this.colMax.OptionsColumn.AllowEdit = false;
			this.colMax.Visible = true;
			this.colMax.VisibleIndex = 8;
			this.colMax.Width = 83;
			// 
			// gridColumn8
			// 
			this.gridColumn8.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.gridColumn8.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.gridColumn8.AppearanceHeader.Options.UseBackColor = true;
			this.gridColumn8.AppearanceHeader.Options.UseFont = true;
			this.gridColumn8.AppearanceHeader.Options.UseTextOptions = true;
			this.gridColumn8.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.gridColumn8.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.gridColumn8.Caption = "Đơn vị tính";
			this.gridColumn8.FieldName = "Unit";
			this.gridColumn8.Name = "gridColumn8";
			this.gridColumn8.OptionsColumn.AllowEdit = false;
			this.gridColumn8.Visible = true;
			this.gridColumn8.VisibleIndex = 5;
			this.gridColumn8.Width = 72;
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
			this.colProductStepID.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colProductStepID.AppearanceHeader.Options.UseBackColor = true;
			this.colProductStepID.Caption = "ProductStepID";
			this.colProductStepID.FieldName = "ProductStepID";
			this.colProductStepID.Name = "colProductStepID";
			this.colProductStepID.OptionsColumn.AllowEdit = false;
			// 
			// colIsHidden
			// 
			this.colIsHidden.AppearanceCell.Options.UseTextOptions = true;
			this.colIsHidden.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colIsHidden.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colIsHidden.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colIsHidden.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colIsHidden.AppearanceHeader.Options.UseBackColor = true;
			this.colIsHidden.AppearanceHeader.Options.UseFont = true;
			this.colIsHidden.AppearanceHeader.Options.UseTextOptions = true;
			this.colIsHidden.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colIsHidden.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colIsHidden.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colIsHidden.Caption = "Ẩn";
			this.colIsHidden.ColumnEdit = this.repositoryItemCheckEdit1;
			this.colIsHidden.FieldName = "IsHidden";
			this.colIsHidden.Name = "colIsHidden";
			this.colIsHidden.OptionsColumn.AllowEdit = false;
			this.colIsHidden.UnboundType = DevExpress.Data.UnboundColumnType.Boolean;
			this.colIsHidden.Visible = true;
			this.colIsHidden.VisibleIndex = 2;
			// 
			// repositoryItemCheckEdit1
			// 
			this.repositoryItemCheckEdit1.AutoHeight = false;
			this.repositoryItemCheckEdit1.Caption = "Check";
			this.repositoryItemCheckEdit1.Name = "repositoryItemCheckEdit1";
			// 
			// colCheckValueTypes
			// 
			this.colCheckValueTypes.AppearanceCell.Options.UseTextOptions = true;
			this.colCheckValueTypes.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colCheckValueTypes.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colCheckValueTypes.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colCheckValueTypes.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colCheckValueTypes.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colCheckValueTypes.AppearanceHeader.Options.UseBackColor = true;
			this.colCheckValueTypes.AppearanceHeader.Options.UseFont = true;
			this.colCheckValueTypes.AppearanceHeader.Options.UseTextOptions = true;
			this.colCheckValueTypes.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colCheckValueTypes.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colCheckValueTypes.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colCheckValueTypes.Caption = "Loại giá trị";
			this.colCheckValueTypes.FieldName = "CheckValueTypes";
			this.colCheckValueTypes.Name = "colCheckValueTypes";
			this.colCheckValueTypes.Visible = true;
			this.colCheckValueTypes.VisibleIndex = 4;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.label1.Location = new System.Drawing.Point(5, 33);
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
			this.splitContainer1.Panel1.Controls.Add(this.toolStrip1);
			this.splitContainer1.Panel1.Controls.Add(this.grdData);
			// 
			// splitContainer1.Panel2
			// 
			this.splitContainer1.Panel2.Controls.Add(this.chkAll);
			this.splitContainer1.Panel2.Controls.Add(this.mnuMenu);
			this.splitContainer1.Panel2.Controls.Add(this.grdDetail);
			this.splitContainer1.Panel2.Controls.Add(this.label1);
			this.splitContainer1.Size = new System.Drawing.Size(1167, 683);
			this.splitContainer1.SplitterDistance = 350;
			this.splitContainer1.TabIndex = 29;
			// 
			// toolStrip1
			// 
			this.toolStrip1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.toolStrip1.AutoSize = false;
			this.toolStrip1.BackColor = System.Drawing.Color.WhiteSmoke;
			this.toolStrip1.Dock = System.Windows.Forms.DockStyle.None;
			this.toolStrip1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.toolStrip1.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden;
			this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.btnDeleteStep,
            this.btnEditStep,
            this.btnAddStep,
            this.btnRestartCD});
			this.toolStrip1.LayoutStyle = System.Windows.Forms.ToolStripLayoutStyle.Flow;
			this.toolStrip1.Location = new System.Drawing.Point(34, 8);
			this.toolStrip1.Name = "toolStrip1";
			this.toolStrip1.RenderMode = System.Windows.Forms.ToolStripRenderMode.Professional;
			this.toolStrip1.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
			this.toolStrip1.Size = new System.Drawing.Size(311, 37);
			this.toolStrip1.TabIndex = 30;
			this.toolStrip1.Text = "toolStrip2";
			// 
			// btnDeleteStep
			// 
			this.btnDeleteStep.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnDeleteStep.Image = ((System.Drawing.Image)(resources.GetObject("btnDeleteStep.Image")));
			this.btnDeleteStep.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnDeleteStep.Name = "btnDeleteStep";
			this.btnDeleteStep.Size = new System.Drawing.Size(32, 33);
			this.btnDeleteStep.Tag = "frmProductWorking_DeleteCD";
			this.btnDeleteStep.Text = "Xóa";
			this.btnDeleteStep.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnDeleteStep.Click += new System.EventHandler(this.btnDeleteStep_Click);
			// 
			// btnEditStep
			// 
			this.btnEditStep.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnEditStep.Image = ((System.Drawing.Image)(resources.GetObject("btnEditStep.Image")));
			this.btnEditStep.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnEditStep.Name = "btnEditStep";
			this.btnEditStep.Size = new System.Drawing.Size(33, 33);
			this.btnEditStep.Tag = "frmProductWorking_EditCD";
			this.btnEditStep.Text = "Sửa";
			this.btnEditStep.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnEditStep.Click += new System.EventHandler(this.btnEditStep_Click);
			// 
			// btnAddStep
			// 
			this.btnAddStep.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnAddStep.Image = ((System.Drawing.Image)(resources.GetObject("btnAddStep.Image")));
			this.btnAddStep.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnAddStep.Name = "btnAddStep";
			this.btnAddStep.Size = new System.Drawing.Size(43, 33);
			this.btnAddStep.Tag = "frmProductWorking_AddCD";
			this.btnAddStep.Text = "Thêm";
			this.btnAddStep.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnAddStep.Click += new System.EventHandler(this.btnAddStep_Click);
			// 
			// btnRestartCD
			// 
			this.btnRestartCD.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnRestartCD.Image = ((System.Drawing.Image)(resources.GetObject("btnRestartCD.Image")));
			this.btnRestartCD.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnRestartCD.Name = "btnRestartCD";
			this.btnRestartCD.Size = new System.Drawing.Size(59, 33);
			this.btnRestartCD.Tag = "";
			this.btnRestartCD.Text = "Làm mới";
			this.btnRestartCD.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnRestartCD.Click += new System.EventHandler(this.btnRestartCD_Click);
			// 
			// chkAll
			// 
			this.chkAll.AutoSize = true;
			this.chkAll.Location = new System.Drawing.Point(9, 10);
			this.chkAll.Name = "chkAll";
			this.chkAll.Size = new System.Drawing.Size(92, 17);
			this.chkAll.TabIndex = 30;
			this.chkAll.Text = "Hiển thị tất cả";
			this.chkAll.UseVisualStyleBackColor = true;
			this.chkAll.CheckedChanged += new System.EventHandler(this.chkAll_CheckedChanged);
			// 
			// mnuMenu
			// 
			this.mnuMenu.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.mnuMenu.AutoSize = false;
			this.mnuMenu.BackColor = System.Drawing.Color.WhiteSmoke;
			this.mnuMenu.Dock = System.Windows.Forms.DockStyle.None;
			this.mnuMenu.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.mnuMenu.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden;
			this.mnuMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.btnDelete,
            this.btnEdit,
            this.btnNew,
            this.toolStripSeparator1,
            this.btnSave,
            this.toolStripButton1});
			this.mnuMenu.LayoutStyle = System.Windows.Forms.ToolStripLayoutStyle.Flow;
			this.mnuMenu.Location = new System.Drawing.Point(438, 9);
			this.mnuMenu.Name = "mnuMenu";
			this.mnuMenu.RenderMode = System.Windows.Forms.ToolStripRenderMode.Professional;
			this.mnuMenu.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
			this.mnuMenu.Size = new System.Drawing.Size(365, 37);
			this.mnuMenu.TabIndex = 29;
			this.mnuMenu.Text = "toolStrip2";
			// 
			// btnDelete
			// 
			this.btnDelete.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnDelete.Image = ((System.Drawing.Image)(resources.GetObject("btnDelete.Image")));
			this.btnDelete.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnDelete.Name = "btnDelete";
			this.btnDelete.Size = new System.Drawing.Size(32, 33);
			this.btnDelete.Tag = "frmProductWorking_DeleteDS";
			this.btnDelete.Text = "Xóa";
			this.btnDelete.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnDelete.Click += new System.EventHandler(this.btnDelete_Click);
			// 
			// btnEdit
			// 
			this.btnEdit.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnEdit.Image = ((System.Drawing.Image)(resources.GetObject("btnEdit.Image")));
			this.btnEdit.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnEdit.Name = "btnEdit";
			this.btnEdit.Size = new System.Drawing.Size(33, 33);
			this.btnEdit.Tag = "frmProductWorking_EditDS";
			this.btnEdit.Text = "Sửa";
			this.btnEdit.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnEdit.Click += new System.EventHandler(this.btnEdit_Click);
			// 
			// btnNew
			// 
			this.btnNew.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnNew.Image = ((System.Drawing.Image)(resources.GetObject("btnNew.Image")));
			this.btnNew.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnNew.Name = "btnNew";
			this.btnNew.Size = new System.Drawing.Size(43, 33);
			this.btnNew.Tag = "frmProductWorking_AddDS";
			this.btnNew.Text = "Thêm";
			this.btnNew.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnNew.Click += new System.EventHandler(this.btnNew_Click);
			// 
			// toolStripSeparator1
			// 
			this.toolStripSeparator1.Name = "toolStripSeparator1";
			this.toolStripSeparator1.Size = new System.Drawing.Size(6, 23);
			// 
			// btnSave
			// 
			this.btnSave.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnSave.Image = ((System.Drawing.Image)(resources.GetObject("btnSave.Image")));
			this.btnSave.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnSave.Name = "btnSave";
			this.btnSave.Size = new System.Drawing.Size(30, 33);
			this.btnSave.Tag = "frmProductWorking_AddDS";
			this.btnSave.Text = "Cất";
			this.btnSave.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnSave.Visible = false;
			this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
			// 
			// toolStripButton1
			// 
			this.toolStripButton1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.toolStripButton1.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButton1.Image")));
			this.toolStripButton1.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.toolStripButton1.Name = "toolStripButton1";
			this.toolStripButton1.Size = new System.Drawing.Size(59, 33);
			this.toolStripButton1.Tag = "";
			this.toolStripButton1.Text = "Làm mới";
			this.toolStripButton1.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.toolStripButton1.Click += new System.EventHandler(this.toolStripButton1_Click);
			// 
			// frmProductWorking
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(1167, 683);
			this.Controls.Add(this.splitContainer1);
			this.Name = "frmProductWorking";
			this.Text = "DANH SÁCH MỤC KIỂM TRA THEO SẢN PHẨM";
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
			this.splitContainer1.Panel2.PerformLayout();
			((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
			this.splitContainer1.ResumeLayout(false);
			this.toolStrip1.ResumeLayout(false);
			this.toolStrip1.PerformLayout();
			this.mnuMenu.ResumeLayout(false);
			this.mnuMenu.PerformLayout();
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
        private DevExpress.XtraGrid.Columns.GridColumn gridColumn4;
        private DevExpress.XtraGrid.Columns.GridColumn colValue;
        private DevExpress.XtraGrid.Columns.GridColumn colMin;
        private DevExpress.XtraGrid.Columns.GridColumn colMax;
        private DevExpress.XtraGrid.Columns.GridColumn gridColumn8;
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
        private System.Windows.Forms.CheckBox chkAll;
        private DevExpress.XtraEditors.Repository.RepositoryItemTextEdit repositoryItemTextEdit1;
        private DevExpress.XtraGrid.Columns.GridColumn colProductStepID;
        private DevExpress.XtraGrid.Columns.GridColumn colIsHidden;
        private DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit repositoryItemCheckEdit1;
        private System.Windows.Forms.ToolStripButton btnRestartCD;
        private System.Windows.Forms.ToolStripButton toolStripButton1;
        private DevExpress.XtraGrid.Columns.GridColumn colCheckValueTypes;
    }
}