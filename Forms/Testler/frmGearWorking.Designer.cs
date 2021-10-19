namespace BMS
{
    partial class frmGearWorking
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
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmGearWorking));
			this.splitContainer1 = new System.Windows.Forms.SplitContainer();
			this.grdData = new DevExpress.XtraGrid.GridControl();
			this.grvData = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.colID = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colHYP = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colGear1 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colGear2 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colGear3 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colLap = new DevExpress.XtraGrid.Columns.GridColumn();
			this.mnuMenu = new System.Windows.Forms.ToolStrip();
			this.btnDelete = new System.Windows.Forms.ToolStripButton();
			this.btnEdit = new System.Windows.Forms.ToolStripButton();
			this.btnNew = new System.Windows.Forms.ToolStripButton();
			this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
			this.btnSave = new System.Windows.Forms.ToolStripButton();
			this.grdDetail = new DevExpress.XtraGrid.GridControl();
			this.grvDetail = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.colGearID = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colID2 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colWorkingName = new DevExpress.XtraGrid.Columns.GridColumn();
			this.repositoryItemMemoEdit1 = new DevExpress.XtraEditors.Repository.RepositoryItemMemoEdit();
			this.colValueTypeName = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colValue = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colMin = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colMax = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colDefaultValue = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colUnit = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colSortOrder = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colTanSuat = new DevExpress.XtraGrid.Columns.GridColumn();
			((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
			this.splitContainer1.Panel1.SuspendLayout();
			this.splitContainer1.Panel2.SuspendLayout();
			this.splitContainer1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.grdData)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grvData)).BeginInit();
			this.mnuMenu.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.grdDetail)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grvDetail)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemMemoEdit1)).BeginInit();
			this.SuspendLayout();
			// 
			// splitContainer1
			// 
			this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.splitContainer1.Location = new System.Drawing.Point(0, 0);
			this.splitContainer1.Name = "splitContainer1";
			// 
			// splitContainer1.Panel1
			// 
			this.splitContainer1.Panel1.Controls.Add(this.grdData);
			// 
			// splitContainer1.Panel2
			// 
			this.splitContainer1.Panel2.Controls.Add(this.mnuMenu);
			this.splitContainer1.Panel2.Controls.Add(this.grdDetail);
			this.splitContainer1.Size = new System.Drawing.Size(1147, 611);
			this.splitContainer1.SplitterDistance = 357;
			this.splitContainer1.TabIndex = 0;
			// 
			// grdData
			// 
			this.grdData.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.grdData.Location = new System.Drawing.Point(3, 17);
			this.grdData.MainView = this.grvData;
			this.grdData.Name = "grdData";
			this.grdData.Size = new System.Drawing.Size(351, 582);
			this.grdData.TabIndex = 0;
			this.grdData.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.grvData});
			// 
			// grvData
			// 
			this.grvData.Appearance.FocusedRow.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.grvData.Appearance.FocusedRow.Options.UseFont = true;
			this.grvData.Appearance.HeaderPanel.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.grvData.Appearance.HeaderPanel.Options.UseFont = true;
			this.grvData.Appearance.HideSelectionRow.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.grvData.Appearance.HideSelectionRow.Options.UseFont = true;
			this.grvData.Appearance.SelectedRow.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.grvData.Appearance.SelectedRow.Options.UseFont = true;
			this.grvData.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
            this.colID,
            this.colHYP,
            this.colGear1,
            this.colGear2,
            this.colGear3,
            this.colLap});
			this.grvData.GridControl = this.grdData;
			this.grvData.Name = "grvData";
			this.grvData.OptionsBehavior.AutoExpandAllGroups = true;
			this.grvData.OptionsBehavior.Editable = false;
			this.grvData.OptionsCustomization.AllowRowSizing = true;
			this.grvData.OptionsFind.AlwaysVisible = true;
			this.grvData.OptionsSelection.MultiSelect = true;
			this.grvData.OptionsView.ColumnAutoWidth = false;
			this.grvData.OptionsView.ShowGroupPanel = false;
			this.grvData.FocusedRowChanged += new DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventHandler(this.grvData_FocusedRowChanged);
			// 
			// colID
			// 
			this.colID.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colID.AppearanceHeader.Options.UseBackColor = true;
			this.colID.AppearanceHeader.Options.UseTextOptions = true;
			this.colID.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colID.Caption = "ID";
			this.colID.FieldName = "ID";
			this.colID.Name = "colID";
			// 
			// colHYP
			// 
			this.colHYP.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colHYP.AppearanceHeader.Options.UseBackColor = true;
			this.colHYP.AppearanceHeader.Options.UseTextOptions = true;
			this.colHYP.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colHYP.Caption = "Mã HYP";
			this.colHYP.FieldName = "HYP";
			this.colHYP.Name = "colHYP";
			this.colHYP.Visible = true;
			this.colHYP.VisibleIndex = 0;
			this.colHYP.Width = 93;
			// 
			// colGear1
			// 
			this.colGear1.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colGear1.AppearanceHeader.Options.UseBackColor = true;
			this.colGear1.AppearanceHeader.Options.UseTextOptions = true;
			this.colGear1.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colGear1.Caption = "Mã Gear 1";
			this.colGear1.FieldName = "Gear1";
			this.colGear1.Name = "colGear1";
			this.colGear1.Visible = true;
			this.colGear1.VisibleIndex = 1;
			this.colGear1.Width = 110;
			// 
			// colGear2
			// 
			this.colGear2.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colGear2.AppearanceHeader.Options.UseBackColor = true;
			this.colGear2.AppearanceHeader.Options.UseTextOptions = true;
			this.colGear2.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colGear2.Caption = "Mã Gear 2";
			this.colGear2.FieldName = "Gear2";
			this.colGear2.Name = "colGear2";
			this.colGear2.Visible = true;
			this.colGear2.VisibleIndex = 2;
			this.colGear2.Width = 99;
			// 
			// colGear3
			// 
			this.colGear3.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colGear3.AppearanceHeader.Options.UseBackColor = true;
			this.colGear3.AppearanceHeader.Options.UseTextOptions = true;
			this.colGear3.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colGear3.Caption = "Mã Gear 3";
			this.colGear3.FieldName = "Gear3";
			this.colGear3.Name = "colGear3";
			this.colGear3.Visible = true;
			this.colGear3.VisibleIndex = 3;
			this.colGear3.Width = 100;
			// 
			// colLap
			// 
			this.colLap.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colLap.AppearanceHeader.Options.UseBackColor = true;
			this.colLap.Caption = "Mã Lap";
			this.colLap.FieldName = "Lap";
			this.colLap.Name = "colLap";
			this.colLap.Visible = true;
			this.colLap.VisibleIndex = 4;
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
            this.btnSave});
			this.mnuMenu.LayoutStyle = System.Windows.Forms.ToolStripLayoutStyle.Flow;
			this.mnuMenu.Location = new System.Drawing.Point(418, 9);
			this.mnuMenu.Name = "mnuMenu";
			this.mnuMenu.RenderMode = System.Windows.Forms.ToolStripRenderMode.Professional;
			this.mnuMenu.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
			this.mnuMenu.Size = new System.Drawing.Size(365, 37);
			this.mnuMenu.TabIndex = 30;
			this.mnuMenu.Text = "toolStrip2";
			// 
			// btnDelete
			// 
			this.btnDelete.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnDelete.Image = ((System.Drawing.Image)(resources.GetObject("btnDelete.Image")));
			this.btnDelete.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnDelete.Name = "btnDelete";
			this.btnDelete.Size = new System.Drawing.Size(32, 33);
			this.btnDelete.Tag = "";
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
			this.btnEdit.Tag = "";
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
			this.btnNew.Tag = "";
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
			this.btnSave.Text = "Cất";
			this.btnSave.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
			// 
			// grdDetail
			// 
			this.grdDetail.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.grdDetail.Location = new System.Drawing.Point(3, 59);
			this.grdDetail.MainView = this.grvDetail;
			this.grdDetail.Name = "grdDetail";
			this.grdDetail.RepositoryItems.AddRange(new DevExpress.XtraEditors.Repository.RepositoryItem[] {
            this.repositoryItemMemoEdit1});
			this.grdDetail.Size = new System.Drawing.Size(780, 549);
			this.grdDetail.TabIndex = 26;
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
            this.colGearID,
            this.colID2,
            this.colWorkingName,
            this.colValueTypeName,
            this.colValue,
            this.colMin,
            this.colMax,
            this.colDefaultValue,
            this.colUnit,
            this.colSortOrder,
            this.colTanSuat});
			this.grvDetail.GridControl = this.grdDetail;
			this.grvDetail.Name = "grvDetail";
			this.grvDetail.OptionsBehavior.Editable = false;
			this.grvDetail.OptionsCustomization.AllowRowSizing = true;
			this.grvDetail.OptionsFind.AlwaysVisible = true;
			this.grvDetail.OptionsView.ColumnAutoWidth = false;
			this.grvDetail.OptionsView.RowAutoHeight = true;
			this.grvDetail.OptionsView.ShowGroupPanel = false;
			this.grvDetail.FocusedRowChanged += new DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventHandler(this.grvDetail_FocusedRowChanged);
			// 
			// colGearID
			// 
			this.colGearID.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colGearID.AppearanceHeader.Options.UseBackColor = true;
			this.colGearID.Caption = "Gear ID";
			this.colGearID.FieldName = "GearID";
			this.colGearID.Name = "colGearID";
			// 
			// colID2
			// 
			this.colID2.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colID2.AppearanceHeader.Options.UseBackColor = true;
			this.colID2.Caption = "ID";
			this.colID2.FieldName = "ID";
			this.colID2.Name = "colID2";
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
			this.colWorkingName.OptionsColumn.AllowFocus = false;
			this.colWorkingName.OptionsColumn.AllowSize = false;
			this.colWorkingName.Visible = true;
			this.colWorkingName.VisibleIndex = 1;
			this.colWorkingName.Width = 224;
			// 
			// repositoryItemMemoEdit1
			// 
			this.repositoryItemMemoEdit1.Appearance.Options.UseTextOptions = true;
			this.repositoryItemMemoEdit1.Appearance.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.repositoryItemMemoEdit1.Name = "repositoryItemMemoEdit1";
			// 
			// colValueTypeName
			// 
			this.colValueTypeName.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colValueTypeName.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colValueTypeName.AppearanceHeader.Options.UseBackColor = true;
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
			this.colValueTypeName.Width = 84;
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
			this.colValue.VisibleIndex = 4;
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
			this.colMin.FieldName = "MinValue";
			this.colMin.Name = "colMin";
			this.colMin.OptionsColumn.AllowEdit = false;
			this.colMin.Visible = true;
			this.colMin.VisibleIndex = 5;
			this.colMin.Width = 106;
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
			this.colMax.FieldName = "MaxValue";
			this.colMax.Name = "colMax";
			this.colMax.OptionsColumn.AllowEdit = false;
			this.colMax.Visible = true;
			this.colMax.VisibleIndex = 6;
			this.colMax.Width = 123;
			// 
			// colDefaultValue
			// 
			this.colDefaultValue.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colDefaultValue.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colDefaultValue.AppearanceHeader.Options.UseBackColor = true;
			this.colDefaultValue.AppearanceHeader.Options.UseFont = true;
			this.colDefaultValue.AppearanceHeader.Options.UseTextOptions = true;
			this.colDefaultValue.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colDefaultValue.Caption = "Giá trị mặc định";
			this.colDefaultValue.FieldName = "DefaultValue";
			this.colDefaultValue.Name = "colDefaultValue";
			this.colDefaultValue.Visible = true;
			this.colDefaultValue.VisibleIndex = 7;
			this.colDefaultValue.Width = 120;
			// 
			// colUnit
			// 
			this.colUnit.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colUnit.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colUnit.AppearanceHeader.Options.UseBackColor = true;
			this.colUnit.AppearanceHeader.Options.UseFont = true;
			this.colUnit.AppearanceHeader.Options.UseTextOptions = true;
			this.colUnit.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colUnit.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colUnit.Caption = "Đơn vị tính";
			this.colUnit.FieldName = "Unit";
			this.colUnit.Name = "colUnit";
			this.colUnit.OptionsColumn.AllowEdit = false;
			this.colUnit.Visible = true;
			this.colUnit.VisibleIndex = 3;
			this.colUnit.Width = 83;
			// 
			// colSortOrder
			// 
			this.colSortOrder.AppearanceCell.Options.UseTextOptions = true;
			this.colSortOrder.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colSortOrder.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colSortOrder.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colSortOrder.AppearanceHeader.Options.UseBackColor = true;
			this.colSortOrder.AppearanceHeader.Options.UseFont = true;
			this.colSortOrder.AppearanceHeader.Options.UseTextOptions = true;
			this.colSortOrder.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colSortOrder.Caption = "STT";
			this.colSortOrder.FieldName = "SortOrder";
			this.colSortOrder.Name = "colSortOrder";
			this.colSortOrder.Visible = true;
			this.colSortOrder.VisibleIndex = 0;
			this.colSortOrder.Width = 51;
			// 
			// colTanSuat
			// 
			this.colTanSuat.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colTanSuat.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colTanSuat.AppearanceHeader.Options.UseBackColor = true;
			this.colTanSuat.AppearanceHeader.Options.UseFont = true;
			this.colTanSuat.AppearanceHeader.Options.UseTextOptions = true;
			this.colTanSuat.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colTanSuat.Caption = "Tần suất";
			this.colTanSuat.FieldName = "TanSuat";
			this.colTanSuat.Name = "colTanSuat";
			this.colTanSuat.Visible = true;
			this.colTanSuat.VisibleIndex = 8;
			this.colTanSuat.Width = 100;
			// 
			// frmGearWorking
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(1147, 611);
			this.Controls.Add(this.splitContainer1);
			this.Name = "frmGearWorking";
			this.Text = "DANH MỤC KIỂM TRA SẢN PHẨM";
			this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
			this.Load += new System.EventHandler(this.frmGearWorking_Load);
			this.splitContainer1.Panel1.ResumeLayout(false);
			this.splitContainer1.Panel2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
			this.splitContainer1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.grdData)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grvData)).EndInit();
			this.mnuMenu.ResumeLayout(false);
			this.mnuMenu.PerformLayout();
			((System.ComponentModel.ISupportInitialize)(this.grdDetail)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grvDetail)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemMemoEdit1)).EndInit();
			this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.SplitContainer splitContainer1;
        private DevExpress.XtraGrid.GridControl grdData;
        private DevExpress.XtraGrid.Views.Grid.GridView grvData;
        private DevExpress.XtraGrid.Columns.GridColumn colID;
        private DevExpress.XtraGrid.Columns.GridColumn colHYP;
        private DevExpress.XtraGrid.Columns.GridColumn colGear1;
        private DevExpress.XtraGrid.Columns.GridColumn colGear2;
        private DevExpress.XtraGrid.Columns.GridColumn colGear3;
        private DevExpress.XtraGrid.Columns.GridColumn colLap;
        private DevExpress.XtraGrid.GridControl grdDetail;
        private DevExpress.XtraGrid.Views.Grid.GridView grvDetail;
        private DevExpress.XtraGrid.Columns.GridColumn colGearID;
        private DevExpress.XtraGrid.Columns.GridColumn colWorkingName;
        private DevExpress.XtraEditors.Repository.RepositoryItemMemoEdit repositoryItemMemoEdit1;
        private DevExpress.XtraGrid.Columns.GridColumn colValueTypeName;
        private DevExpress.XtraGrid.Columns.GridColumn colValue;
        private DevExpress.XtraGrid.Columns.GridColumn colMin;
        private DevExpress.XtraGrid.Columns.GridColumn colMax;
        private DevExpress.XtraGrid.Columns.GridColumn colUnit;
        private DevExpress.XtraGrid.Columns.GridColumn colSortOrder;
        private System.Windows.Forms.ToolStrip mnuMenu;
        private System.Windows.Forms.ToolStripButton btnDelete;
        private System.Windows.Forms.ToolStripButton btnEdit;
        private System.Windows.Forms.ToolStripButton btnNew;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripButton btnSave;
        private DevExpress.XtraGrid.Columns.GridColumn colDefaultValue;
        private DevExpress.XtraGrid.Columns.GridColumn colTanSuat;
        private DevExpress.XtraGrid.Columns.GridColumn colID2;
    }
}