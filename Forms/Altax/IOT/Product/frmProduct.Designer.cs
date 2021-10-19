namespace BMS
{
    partial class frmProduct
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
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmProduct));
			this.colS = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colID = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colGroup = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colCode = new DevExpress.XtraGrid.Columns.GridColumn();
			this.repositoryItemMemoEdit1 = new DevExpress.XtraEditors.Repository.RepositoryItemMemoEdit();
			this.colName = new DevExpress.XtraGrid.Columns.GridColumn();
			this.grvData = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.colUpdatedDate = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colMotorCode = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colHuongHopCau = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colRatioCode = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colDoDao = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colMURATA = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colCoilCode = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colLoaiMo = new DevExpress.XtraGrid.Columns.GridColumn();
			this.grdData = new DevExpress.XtraGrid.GridControl();
			this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
			this.btnShowWorking = new System.Windows.Forms.ToolStripMenuItem();
			this.repositoryItemCheckEdit1 = new DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit();
			this.repositoryItemCheckEdit2 = new DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit();
			this.btnImportExcel = new System.Windows.Forms.ToolStripButton();
			this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
			this.btnDel = new System.Windows.Forms.ToolStripButton();
			this.toolStripSeparator4 = new System.Windows.Forms.ToolStripSeparator();
			this.btnEdit = new System.Windows.Forms.ToolStripButton();
			this.toolStripSeparator6 = new System.Windows.Forms.ToolStripSeparator();
			this.btnAdd = new System.Windows.Forms.ToolStripButton();
			this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
			this.btnDeleteGroup = new System.Windows.Forms.ToolStripButton();
			this.toolStripSeparator8 = new System.Windows.Forms.ToolStripSeparator();
			this.btnEditGroup = new System.Windows.Forms.ToolStripButton();
			this.toolStripSeparator9 = new System.Windows.Forms.ToolStripSeparator();
			this.btnNewGroup = new System.Windows.Forms.ToolStripButton();
			this.mnuMenu = new System.Windows.Forms.ToolStrip();
			this.btnSungLuc = new System.Windows.Forms.ToolStripButton();
			this.toolStripSeparator5 = new System.Windows.Forms.ToolStripSeparator();
			this.btnExcel = new System.Windows.Forms.ToolStripButton();
			this.toolStripSeparator3 = new System.Windows.Forms.ToolStripSeparator();
			this.btnUpdateStepCode = new System.Windows.Forms.ToolStripButton();
			this.btnSpeedReducer = new System.Windows.Forms.ToolStripButton();
			this.splitContainer1 = new System.Windows.Forms.SplitContainer();
			this.labelControl1 = new DevExpress.XtraEditors.LabelControl();
			this.treeData = new DevExpress.XtraTreeList.TreeList();
			this.colIDTree = new DevExpress.XtraTreeList.Columns.TreeListColumn();
			this.colNameTree = new DevExpress.XtraTreeList.Columns.TreeListColumn();
			this.repositoryItemMemoEdit2 = new DevExpress.XtraEditors.Repository.RepositoryItemMemoEdit();
			this.colCodeTree = new DevExpress.XtraTreeList.Columns.TreeListColumn();
			this.btnFindAll = new DevExpress.XtraEditors.SimpleButton();
			this.txtName = new System.Windows.Forms.TextBox();
			this.label2 = new System.Windows.Forms.Label();
			this.labelControl2 = new DevExpress.XtraEditors.LabelControl();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemMemoEdit1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grvData)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grdData)).BeginInit();
			this.contextMenuStrip1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemCheckEdit1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemCheckEdit2)).BeginInit();
			this.mnuMenu.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
			this.splitContainer1.Panel1.SuspendLayout();
			this.splitContainer1.Panel2.SuspendLayout();
			this.splitContainer1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.treeData)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemMemoEdit2)).BeginInit();
			this.SuspendLayout();
			// 
			// colS
			// 
			this.colS.AppearanceCell.Options.UseTextOptions = true;
			this.colS.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colS.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colS.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colS.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colS.AppearanceHeader.Options.UseBackColor = true;
			this.colS.AppearanceHeader.Options.UseFont = true;
			this.colS.AppearanceHeader.Options.UseTextOptions = true;
			this.colS.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colS.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colS.Caption = "Số pin";
			this.colS.FieldName = "PinNumber";
			this.colS.Name = "colS";
			this.colS.OptionsColumn.AllowEdit = false;
			this.colS.Visible = true;
			this.colS.VisibleIndex = 5;
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
			// colGroup
			// 
			this.colGroup.AppearanceCell.Options.UseTextOptions = true;
			this.colGroup.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colGroup.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colGroup.AppearanceHeader.Options.UseBackColor = true;
			this.colGroup.AppearanceHeader.Options.UseTextOptions = true;
			this.colGroup.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colGroup.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colGroup.Caption = "Nhóm";
			this.colGroup.FieldName = "ProductGroupID";
			this.colGroup.Name = "colGroup";
			this.colGroup.OptionsColumn.AllowEdit = false;
			// 
			// colCode
			// 
			this.colCode.AppearanceCell.Options.UseTextOptions = true;
			this.colCode.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colCode.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colCode.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colCode.AppearanceHeader.Options.UseBackColor = true;
			this.colCode.AppearanceHeader.Options.UseFont = true;
			this.colCode.AppearanceHeader.Options.UseTextOptions = true;
			this.colCode.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colCode.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colCode.Caption = "Mã";
			this.colCode.FieldName = "ProductCode";
			this.colCode.Name = "colCode";
			this.colCode.OptionsColumn.AllowEdit = false;
			this.colCode.Summary.AddRange(new DevExpress.XtraGrid.GridSummaryItem[] {
            new DevExpress.XtraGrid.GridColumnSummaryItem(DevExpress.Data.SummaryItemType.Count)});
			this.colCode.Visible = true;
			this.colCode.VisibleIndex = 0;
			this.colCode.Width = 101;
			// 
			// repositoryItemMemoEdit1
			// 
			this.repositoryItemMemoEdit1.Appearance.Options.UseTextOptions = true;
			this.repositoryItemMemoEdit1.Appearance.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.repositoryItemMemoEdit1.Name = "repositoryItemMemoEdit1";
			// 
			// colName
			// 
			this.colName.AppearanceCell.Options.UseTextOptions = true;
			this.colName.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colName.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colName.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colName.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colName.AppearanceHeader.Options.UseBackColor = true;
			this.colName.AppearanceHeader.Options.UseFont = true;
			this.colName.AppearanceHeader.Options.UseTextOptions = true;
			this.colName.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colName.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colName.Caption = "Tên";
			this.colName.ColumnEdit = this.repositoryItemMemoEdit1;
			this.colName.FieldName = "ProductName";
			this.colName.Name = "colName";
			this.colName.OptionsColumn.AllowEdit = false;
			this.colName.OptionsColumn.AllowSize = false;
			this.colName.Visible = true;
			this.colName.VisibleIndex = 1;
			this.colName.Width = 251;
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
            this.colName,
            this.colCode,
            this.colGroup,
            this.colID,
            this.colS,
            this.colUpdatedDate,
            this.colMotorCode,
            this.colHuongHopCau,
            this.colRatioCode,
            this.colDoDao,
            this.colMURATA,
            this.colCoilCode,
            this.colLoaiMo});
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
			this.grvData.OptionsView.ShowAutoFilterRow = true;
			this.grvData.OptionsView.ShowFooter = true;
			this.grvData.OptionsView.ShowGroupPanel = false;
			this.grvData.DoubleClick += new System.EventHandler(this.grvData_DoubleClick);
			// 
			// colUpdatedDate
			// 
			this.colUpdatedDate.AppearanceCell.Options.UseTextOptions = true;
			this.colUpdatedDate.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colUpdatedDate.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colUpdatedDate.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colUpdatedDate.AppearanceHeader.Options.UseBackColor = true;
			this.colUpdatedDate.AppearanceHeader.Options.UseFont = true;
			this.colUpdatedDate.AppearanceHeader.Options.UseTextOptions = true;
			this.colUpdatedDate.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colUpdatedDate.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colUpdatedDate.Caption = "Ngày sửa";
			this.colUpdatedDate.FieldName = "UpdatedDate";
			this.colUpdatedDate.Name = "colUpdatedDate";
			this.colUpdatedDate.OptionsColumn.AllowEdit = false;
			// 
			// colMotorCode
			// 
			this.colMotorCode.AppearanceCell.Options.UseTextOptions = true;
			this.colMotorCode.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colMotorCode.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colMotorCode.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colMotorCode.AppearanceHeader.Options.UseBackColor = true;
			this.colMotorCode.AppearanceHeader.Options.UseFont = true;
			this.colMotorCode.AppearanceHeader.Options.UseTextOptions = true;
			this.colMotorCode.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colMotorCode.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colMotorCode.Caption = "MÃ MOTOR";
			this.colMotorCode.FieldName = "MotorCode";
			this.colMotorCode.Name = "colMotorCode";
			this.colMotorCode.OptionsColumn.AllowEdit = false;
			this.colMotorCode.Visible = true;
			this.colMotorCode.VisibleIndex = 7;
			// 
			// colHuongHopCau
			// 
			this.colHuongHopCau.AppearanceCell.Options.UseTextOptions = true;
			this.colHuongHopCau.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colHuongHopCau.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colHuongHopCau.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colHuongHopCau.AppearanceHeader.Options.UseBackColor = true;
			this.colHuongHopCau.AppearanceHeader.Options.UseFont = true;
			this.colHuongHopCau.AppearanceHeader.Options.UseTextOptions = true;
			this.colHuongHopCau.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colHuongHopCau.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colHuongHopCau.Caption = "Hướng hộp cầu";
			this.colHuongHopCau.FieldName = "HuongHopCau";
			this.colHuongHopCau.Name = "colHuongHopCau";
			this.colHuongHopCau.OptionsColumn.AllowEdit = false;
			this.colHuongHopCau.Width = 169;
			// 
			// colRatioCode
			// 
			this.colRatioCode.AppearanceCell.Options.UseTextOptions = true;
			this.colRatioCode.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colRatioCode.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colRatioCode.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colRatioCode.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colRatioCode.AppearanceHeader.Options.UseBackColor = true;
			this.colRatioCode.AppearanceHeader.Options.UseFont = true;
			this.colRatioCode.AppearanceHeader.Options.UseTextOptions = true;
			this.colRatioCode.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colRatioCode.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colRatioCode.Caption = "Ratio";
			this.colRatioCode.FieldName = "RatioCode";
			this.colRatioCode.Name = "colRatioCode";
			this.colRatioCode.OptionsColumn.AllowEdit = false;
			this.colRatioCode.OptionsColumn.AllowSize = false;
			this.colRatioCode.Visible = true;
			this.colRatioCode.VisibleIndex = 2;
			this.colRatioCode.Width = 92;
			// 
			// colDoDao
			// 
			this.colDoDao.AppearanceCell.Options.UseTextOptions = true;
			this.colDoDao.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colDoDao.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colDoDao.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colDoDao.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colDoDao.AppearanceHeader.Options.UseBackColor = true;
			this.colDoDao.AppearanceHeader.Options.UseFont = true;
			this.colDoDao.AppearanceHeader.Options.UseTextOptions = true;
			this.colDoDao.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colDoDao.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colDoDao.Caption = "Độ dảo";
			this.colDoDao.FieldName = "DoDao";
			this.colDoDao.Name = "colDoDao";
			this.colDoDao.OptionsColumn.AllowEdit = false;
			this.colDoDao.OptionsColumn.AllowSize = false;
			this.colDoDao.Visible = true;
			this.colDoDao.VisibleIndex = 3;
			this.colDoDao.Width = 67;
			// 
			// colMURATA
			// 
			this.colMURATA.AppearanceCell.Options.UseTextOptions = true;
			this.colMURATA.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colMURATA.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colMURATA.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colMURATA.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colMURATA.AppearanceHeader.Options.UseBackColor = true;
			this.colMURATA.AppearanceHeader.Options.UseFont = true;
			this.colMURATA.AppearanceHeader.Options.UseTextOptions = true;
			this.colMURATA.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colMURATA.Caption = "MURATA";
			this.colMURATA.FieldName = "MURATA";
			this.colMURATA.Name = "colMURATA";
			this.colMURATA.OptionsColumn.AllowEdit = false;
			this.colMURATA.OptionsColumn.AllowSize = false;
			this.colMURATA.UnboundType = DevExpress.Data.UnboundColumnType.Boolean;
			this.colMURATA.Visible = true;
			this.colMURATA.VisibleIndex = 4;
			this.colMURATA.Width = 70;
			// 
			// colCoilCode
			// 
			this.colCoilCode.AppearanceCell.Options.UseTextOptions = true;
			this.colCoilCode.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colCoilCode.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colCoilCode.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colCoilCode.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colCoilCode.AppearanceHeader.Options.UseBackColor = true;
			this.colCoilCode.AppearanceHeader.Options.UseFont = true;
			this.colCoilCode.AppearanceHeader.Options.UseTextOptions = true;
			this.colCoilCode.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colCoilCode.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colCoilCode.Caption = "Mã Coil";
			this.colCoilCode.FieldName = "CoilCode";
			this.colCoilCode.Name = "colCoilCode";
			this.colCoilCode.OptionsColumn.AllowEdit = false;
			this.colCoilCode.Visible = true;
			this.colCoilCode.VisibleIndex = 6;
			this.colCoilCode.Width = 96;
			// 
			// colLoaiMo
			// 
			this.colLoaiMo.AppearanceCell.Options.UseTextOptions = true;
			this.colLoaiMo.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colLoaiMo.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colLoaiMo.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colLoaiMo.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colLoaiMo.AppearanceHeader.Options.UseBackColor = true;
			this.colLoaiMo.AppearanceHeader.Options.UseFont = true;
			this.colLoaiMo.AppearanceHeader.Options.UseTextOptions = true;
			this.colLoaiMo.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colLoaiMo.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colLoaiMo.Caption = "Loại mỡ";
			this.colLoaiMo.ColumnEdit = this.repositoryItemMemoEdit1;
			this.colLoaiMo.FieldName = "LoaiMo";
			this.colLoaiMo.Name = "colLoaiMo";
			this.colLoaiMo.OptionsColumn.AllowEdit = false;
			this.colLoaiMo.Visible = true;
			this.colLoaiMo.VisibleIndex = 8;
			this.colLoaiMo.Width = 128;
			// 
			// grdData
			// 
			this.grdData.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.grdData.ContextMenuStrip = this.contextMenuStrip1;
			this.grdData.Font = new System.Drawing.Font("Times New Roman", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.grdData.Location = new System.Drawing.Point(6, 35);
			this.grdData.MainView = this.grvData;
			this.grdData.Name = "grdData";
			this.grdData.RepositoryItems.AddRange(new DevExpress.XtraEditors.Repository.RepositoryItem[] {
            this.repositoryItemCheckEdit1,
            this.repositoryItemMemoEdit1,
            this.repositoryItemCheckEdit2});
			this.grdData.Size = new System.Drawing.Size(1000, 588);
			this.grdData.TabIndex = 16;
			this.grdData.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.grvData});
			// 
			// contextMenuStrip1
			// 
			this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.btnShowWorking});
			this.contextMenuStrip1.Name = "contextMenuStrip1";
			this.contextMenuStrip1.Size = new System.Drawing.Size(229, 26);
			// 
			// btnShowWorking
			// 
			this.btnShowWorking.Name = "btnShowWorking";
			this.btnShowWorking.Size = new System.Drawing.Size(228, 22);
			this.btnShowWorking.Tag = "Module_Film";
			this.btnShowWorking.Text = "Xem danh sách mục kiểm tra";
			this.btnShowWorking.Click += new System.EventHandler(this.btnShowWorking_Click);
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
			// btnImportExcel
			// 
			this.btnImportExcel.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnImportExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnImportExcel.Image")));
			this.btnImportExcel.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnImportExcel.Name = "btnImportExcel";
			this.btnImportExcel.Size = new System.Drawing.Size(71, 33);
			this.btnImportExcel.Tag = "frmProduct_ImportExcelProduct";
			this.btnImportExcel.Text = "Nhập Excel";
			this.btnImportExcel.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnImportExcel.Click += new System.EventHandler(this.btnImportExcel_Click);
			// 
			// toolStripSeparator2
			// 
			this.toolStripSeparator2.AutoSize = false;
			this.toolStripSeparator2.Name = "toolStripSeparator2";
			this.toolStripSeparator2.Size = new System.Drawing.Size(6, 40);
			// 
			// btnDel
			// 
			this.btnDel.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnDel.Image = ((System.Drawing.Image)(resources.GetObject("btnDel.Image")));
			this.btnDel.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnDel.Name = "btnDel";
			this.btnDel.Size = new System.Drawing.Size(90, 33);
			this.btnDel.Tag = "frmProduct_DeleteProduct";
			this.btnDel.Text = "Xóa sản phẩm";
			this.btnDel.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnDel.Click += new System.EventHandler(this.btnDel_Click);
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
			this.btnEdit.Size = new System.Drawing.Size(91, 33);
			this.btnEdit.Tag = "frmProduct_EditProduct";
			this.btnEdit.Text = "Sửa sản phẩm";
			this.btnEdit.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnEdit.Click += new System.EventHandler(this.btnEdit_Click);
			// 
			// toolStripSeparator6
			// 
			this.toolStripSeparator6.Name = "toolStripSeparator6";
			this.toolStripSeparator6.Size = new System.Drawing.Size(6, 23);
			// 
			// btnAdd
			// 
			this.btnAdd.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnAdd.Image = ((System.Drawing.Image)(resources.GetObject("btnAdd.Image")));
			this.btnAdd.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnAdd.Name = "btnAdd";
			this.btnAdd.Size = new System.Drawing.Size(90, 33);
			this.btnAdd.Tag = "frmProduct_AddProduct";
			this.btnAdd.Text = "Tạo sản phẩm";
			this.btnAdd.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
			// 
			// toolStripSeparator1
			// 
			this.toolStripSeparator1.AutoSize = false;
			this.toolStripSeparator1.Name = "toolStripSeparator1";
			this.toolStripSeparator1.Size = new System.Drawing.Size(6, 40);
			// 
			// btnDeleteGroup
			// 
			this.btnDeleteGroup.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnDeleteGroup.Image = ((System.Drawing.Image)(resources.GetObject("btnDeleteGroup.Image")));
			this.btnDeleteGroup.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnDeleteGroup.Name = "btnDeleteGroup";
			this.btnDeleteGroup.Size = new System.Drawing.Size(84, 33);
			this.btnDeleteGroup.Tag = "frmProduct_DeleteGroup";
			this.btnDeleteGroup.Text = "Xóa nhóm SP";
			this.btnDeleteGroup.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnDeleteGroup.Click += new System.EventHandler(this.btnDeleteGroup_Click);
			// 
			// toolStripSeparator8
			// 
			this.toolStripSeparator8.Name = "toolStripSeparator8";
			this.toolStripSeparator8.Size = new System.Drawing.Size(6, 23);
			// 
			// btnEditGroup
			// 
			this.btnEditGroup.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnEditGroup.Image = ((System.Drawing.Image)(resources.GetObject("btnEditGroup.Image")));
			this.btnEditGroup.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnEditGroup.Name = "btnEditGroup";
			this.btnEditGroup.Size = new System.Drawing.Size(85, 33);
			this.btnEditGroup.Tag = "frmProduct_EditGroup";
			this.btnEditGroup.Text = "Sửa nhóm SP";
			this.btnEditGroup.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnEditGroup.Click += new System.EventHandler(this.btnEditGroup_Click);
			// 
			// toolStripSeparator9
			// 
			this.toolStripSeparator9.Name = "toolStripSeparator9";
			this.toolStripSeparator9.Size = new System.Drawing.Size(6, 23);
			// 
			// btnNewGroup
			// 
			this.btnNewGroup.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnNewGroup.Image = ((System.Drawing.Image)(resources.GetObject("btnNewGroup.Image")));
			this.btnNewGroup.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnNewGroup.Name = "btnNewGroup";
			this.btnNewGroup.Size = new System.Drawing.Size(84, 33);
			this.btnNewGroup.Tag = "frmProduct_AddGroup";
			this.btnNewGroup.Text = "&Tạo nhóm SP";
			this.btnNewGroup.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnNewGroup.Click += new System.EventHandler(this.btnNewGroup_Click);
			// 
			// mnuMenu
			// 
			this.mnuMenu.AutoSize = false;
			this.mnuMenu.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(234)))), ((int)(((byte)(234)))), ((int)(((byte)(234)))));
			this.mnuMenu.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.mnuMenu.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden;
			this.mnuMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.btnNewGroup,
            this.toolStripSeparator9,
            this.btnEditGroup,
            this.toolStripSeparator8,
            this.btnDeleteGroup,
            this.toolStripSeparator1,
            this.btnAdd,
            this.toolStripSeparator6,
            this.btnEdit,
            this.toolStripSeparator4,
            this.btnDel,
            this.toolStripSeparator2,
            this.btnSungLuc,
            this.toolStripSeparator5,
            this.btnExcel,
            this.toolStripSeparator3,
            this.btnImportExcel,
            this.btnUpdateStepCode,
            this.btnSpeedReducer});
			this.mnuMenu.LayoutStyle = System.Windows.Forms.ToolStripLayoutStyle.Flow;
			this.mnuMenu.Location = new System.Drawing.Point(0, 0);
			this.mnuMenu.Name = "mnuMenu";
			this.mnuMenu.RenderMode = System.Windows.Forms.ToolStripRenderMode.Professional;
			this.mnuMenu.Size = new System.Drawing.Size(1370, 42);
			this.mnuMenu.TabIndex = 20;
			this.mnuMenu.Text = "toolStrip2";
			// 
			// btnSungLuc
			// 
			this.btnSungLuc.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnSungLuc.Image = ((System.Drawing.Image)(resources.GetObject("btnSungLuc.Image")));
			this.btnSungLuc.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnSungLuc.Name = "btnSungLuc";
			this.btnSungLuc.Size = new System.Drawing.Size(62, 33);
			this.btnSungLuc.Tag = "frmProduct_Gun";
			this.btnSungLuc.Text = "Súng Lực";
			this.btnSungLuc.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnSungLuc.Click += new System.EventHandler(this.btnSungLuc_Click);
			// 
			// toolStripSeparator5
			// 
			this.toolStripSeparator5.Name = "toolStripSeparator5";
			this.toolStripSeparator5.Size = new System.Drawing.Size(6, 23);
			// 
			// btnExcel
			// 
			this.btnExcel.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(69, 33);
			this.btnExcel.Tag = "";
			this.btnExcel.Text = "Xuất Excel";
			this.btnExcel.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// toolStripSeparator3
			// 
			this.toolStripSeparator3.Name = "toolStripSeparator3";
			this.toolStripSeparator3.Size = new System.Drawing.Size(6, 23);
			// 
			// btnUpdateStepCode
			// 
			this.btnUpdateStepCode.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnUpdateStepCode.Image = ((System.Drawing.Image)(resources.GetObject("btnUpdateStepCode.Image")));
			this.btnUpdateStepCode.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnUpdateStepCode.Name = "btnUpdateStepCode";
			this.btnUpdateStepCode.Size = new System.Drawing.Size(159, 33);
			this.btnUpdateStepCode.Tag = "frmProduct_UpdateCDH";
			this.btnUpdateStepCode.Text = "Cập nhật lại mã công đoạn";
			this.btnUpdateStepCode.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnUpdateStepCode.Click += new System.EventHandler(this.btnUpdateStepCode_Click);
			// 
			// btnSpeedReducer
			// 
			this.btnSpeedReducer.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnSpeedReducer.Image = ((System.Drawing.Image)(resources.GetObject("btnSpeedReducer.Image")));
			this.btnSpeedReducer.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnSpeedReducer.Name = "btnSpeedReducer";
			this.btnSpeedReducer.Size = new System.Drawing.Size(61, 33);
			this.btnSpeedReducer.Tag = "frmProduct_UpdateCDH";
			this.btnSpeedReducer.Text = "Giảm tốc";
			this.btnSpeedReducer.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnSpeedReducer.Click += new System.EventHandler(this.btnSpeedReducer_Click);
			// 
			// splitContainer1
			// 
			this.splitContainer1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.splitContainer1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(234)))), ((int)(((byte)(234)))), ((int)(((byte)(234)))));
			this.splitContainer1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.splitContainer1.Location = new System.Drawing.Point(0, 45);
			this.splitContainer1.Name = "splitContainer1";
			// 
			// splitContainer1.Panel1
			// 
			this.splitContainer1.Panel1.Controls.Add(this.labelControl1);
			this.splitContainer1.Panel1.Controls.Add(this.treeData);
			// 
			// splitContainer1.Panel2
			// 
			this.splitContainer1.Panel2.Controls.Add(this.btnFindAll);
			this.splitContainer1.Panel2.Controls.Add(this.txtName);
			this.splitContainer1.Panel2.Controls.Add(this.label2);
			this.splitContainer1.Panel2.Controls.Add(this.labelControl2);
			this.splitContainer1.Panel2.Controls.Add(this.grdData);
			this.splitContainer1.Size = new System.Drawing.Size(1370, 630);
			this.splitContainer1.SplitterDistance = 354;
			this.splitContainer1.TabIndex = 21;
			// 
			// labelControl1
			// 
			this.labelControl1.Appearance.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.labelControl1.Appearance.Options.UseFont = true;
			this.labelControl1.Location = new System.Drawing.Point(10, 6);
			this.labelControl1.Name = "labelControl1";
			this.labelControl1.Size = new System.Drawing.Size(151, 13);
			this.labelControl1.TabIndex = 18;
			this.labelControl1.Text = "Danh sách nhóm sản phẩm";
			// 
			// treeData
			// 
			this.treeData.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.treeData.Appearance.FocusedRow.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(128)))), ((int)(((byte)(255)))));
			this.treeData.Appearance.FocusedRow.BackColor2 = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(128)))), ((int)(((byte)(255)))));
			this.treeData.Appearance.FocusedRow.Options.UseBackColor = true;
			this.treeData.Appearance.HideSelectionRow.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(128)))), ((int)(((byte)(255)))));
			this.treeData.Appearance.HideSelectionRow.BackColor2 = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(128)))), ((int)(((byte)(255)))));
			this.treeData.Appearance.HideSelectionRow.Options.UseBackColor = true;
			this.treeData.Columns.AddRange(new DevExpress.XtraTreeList.Columns.TreeListColumn[] {
            this.colIDTree,
            this.colNameTree,
            this.colCodeTree});
			this.treeData.Location = new System.Drawing.Point(5, 23);
			this.treeData.Name = "treeData";
			this.treeData.OptionsBehavior.AllowExpandOnDblClick = false;
			this.treeData.OptionsBehavior.Editable = false;
			this.treeData.OptionsDragAndDrop.DragNodesMode = DevExpress.XtraTreeList.DragNodesMode.Single;
			this.treeData.OptionsView.ShowIndicator = false;
			this.treeData.RepositoryItems.AddRange(new DevExpress.XtraEditors.Repository.RepositoryItem[] {
            this.repositoryItemMemoEdit2});
			this.treeData.Size = new System.Drawing.Size(342, 600);
			this.treeData.TabIndex = 17;
			this.treeData.FocusedNodeChanged += new DevExpress.XtraTreeList.FocusedNodeChangedEventHandler(this.treeData_FocusedNodeChanged);
			// 
			// colIDTree
			// 
			this.colIDTree.Caption = "Mã nhóm";
			this.colIDTree.FieldName = "ID";
			this.colIDTree.Name = "colIDTree";
			// 
			// colNameTree
			// 
			this.colNameTree.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colNameTree.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colNameTree.AppearanceHeader.Options.UseBackColor = true;
			this.colNameTree.AppearanceHeader.Options.UseFont = true;
			this.colNameTree.AppearanceHeader.Options.UseTextOptions = true;
			this.colNameTree.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colNameTree.Caption = "Tên nhóm";
			this.colNameTree.ColumnEdit = this.repositoryItemMemoEdit2;
			this.colNameTree.FieldName = "ProductGroupName";
			this.colNameTree.Name = "colNameTree";
			this.colNameTree.OptionsColumn.AllowEdit = false;
			this.colNameTree.OptionsColumn.AllowFocus = false;
			this.colNameTree.OptionsColumn.AllowMove = false;
			this.colNameTree.OptionsColumn.AllowSort = false;
			this.colNameTree.Visible = true;
			this.colNameTree.VisibleIndex = 1;
			this.colNameTree.Width = 169;
			// 
			// repositoryItemMemoEdit2
			// 
			this.repositoryItemMemoEdit2.Name = "repositoryItemMemoEdit2";
			// 
			// colCodeTree
			// 
			this.colCodeTree.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colCodeTree.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colCodeTree.AppearanceHeader.Options.UseBackColor = true;
			this.colCodeTree.AppearanceHeader.Options.UseFont = true;
			this.colCodeTree.AppearanceHeader.Options.UseTextOptions = true;
			this.colCodeTree.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colCodeTree.Caption = "Mã nhóm";
			this.colCodeTree.FieldName = "ProductGroupCode";
			this.colCodeTree.Name = "colCodeTree";
			this.colCodeTree.OptionsColumn.AllowEdit = false;
			this.colCodeTree.OptionsColumn.AllowFocus = false;
			this.colCodeTree.Visible = true;
			this.colCodeTree.VisibleIndex = 0;
			this.colCodeTree.Width = 134;
			// 
			// btnFindAll
			// 
			this.btnFindAll.Location = new System.Drawing.Point(579, 6);
			this.btnFindAll.Name = "btnFindAll";
			this.btnFindAll.Size = new System.Drawing.Size(75, 23);
			this.btnFindAll.TabIndex = 21;
			this.btnFindAll.Text = "Tìm kiếm";
			this.btnFindAll.Click += new System.EventHandler(this.btnFindAll_Click);
			// 
			// txtName
			// 
			this.txtName.Location = new System.Drawing.Point(277, 7);
			this.txtName.Name = "txtName";
			this.txtName.Size = new System.Drawing.Size(290, 20);
			this.txtName.TabIndex = 20;
			this.txtName.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtName_KeyDown);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(205, 10);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(66, 13);
			this.label2.TabIndex = 19;
			this.label2.Text = "Tên module:";
			// 
			// labelControl2
			// 
			this.labelControl2.Appearance.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.labelControl2.Appearance.Options.UseFont = true;
			this.labelControl2.Location = new System.Drawing.Point(6, 13);
			this.labelControl2.Name = "labelControl2";
			this.labelControl2.Size = new System.Drawing.Size(116, 13);
			this.labelControl2.TabIndex = 18;
			this.labelControl2.Text = "Danh sách sản phẩm";
			// 
			// frmProduct
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(1370, 677);
			this.Controls.Add(this.mnuMenu);
			this.Controls.Add(this.splitContainer1);
			this.Name = "frmProduct";
			this.Text = "DANH SÁCH SẢN PHẨM";
			this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
			this.Load += new System.EventHandler(this.frmProduct_Load);
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemMemoEdit1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grvData)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grdData)).EndInit();
			this.contextMenuStrip1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemCheckEdit1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemCheckEdit2)).EndInit();
			this.mnuMenu.ResumeLayout(false);
			this.mnuMenu.PerformLayout();
			this.splitContainer1.Panel1.ResumeLayout(false);
			this.splitContainer1.Panel1.PerformLayout();
			this.splitContainer1.Panel2.ResumeLayout(false);
			this.splitContainer1.Panel2.PerformLayout();
			((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
			this.splitContainer1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.treeData)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemMemoEdit2)).EndInit();
			this.ResumeLayout(false);

        }

        #endregion

        private DevExpress.XtraGrid.Columns.GridColumn colS;
        private DevExpress.XtraGrid.Columns.GridColumn colID;
        private DevExpress.XtraGrid.Columns.GridColumn colGroup;
        private DevExpress.XtraGrid.Columns.GridColumn colCode;
        private DevExpress.XtraEditors.Repository.RepositoryItemMemoEdit repositoryItemMemoEdit1;
        private DevExpress.XtraGrid.Columns.GridColumn colName;
        private DevExpress.XtraGrid.Views.Grid.GridView grvData;
        private DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit repositoryItemCheckEdit1;
        private DevExpress.XtraGrid.Columns.GridColumn colUpdatedDate;
        private DevExpress.XtraGrid.Columns.GridColumn colMotorCode;
        private DevExpress.XtraGrid.Columns.GridColumn colHuongHopCau;
        private DevExpress.XtraGrid.Columns.GridColumn colRatioCode;
        private DevExpress.XtraGrid.Columns.GridColumn colDoDao;
        private DevExpress.XtraGrid.Columns.GridColumn colMURATA;
        private DevExpress.XtraGrid.Columns.GridColumn colCoilCode;
        private DevExpress.XtraGrid.Columns.GridColumn colLoaiMo;
        private DevExpress.XtraGrid.GridControl grdData;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem btnShowWorking;
        private DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit repositoryItemCheckEdit2;
        private System.Windows.Forms.ToolStripButton btnImportExcel;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
        private System.Windows.Forms.ToolStripButton btnDel;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator4;
        private System.Windows.Forms.ToolStripButton btnEdit;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator6;
        private System.Windows.Forms.ToolStripButton btnAdd;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripButton btnDeleteGroup;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator8;
        private System.Windows.Forms.ToolStripButton btnEditGroup;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator9;
        private System.Windows.Forms.ToolStripButton btnNewGroup;
        private System.Windows.Forms.ToolStrip mnuMenu;
        private System.Windows.Forms.ToolStripButton btnExcel;
        private System.Windows.Forms.SplitContainer splitContainer1;
        private DevExpress.XtraEditors.LabelControl labelControl1;
        private DevExpress.XtraTreeList.TreeList treeData;
        private DevExpress.XtraTreeList.Columns.TreeListColumn colIDTree;
        private DevExpress.XtraTreeList.Columns.TreeListColumn colNameTree;
        private DevExpress.XtraEditors.Repository.RepositoryItemMemoEdit repositoryItemMemoEdit2;
        private DevExpress.XtraTreeList.Columns.TreeListColumn colCodeTree;
        private DevExpress.XtraEditors.SimpleButton btnFindAll;
        private System.Windows.Forms.TextBox txtName;
        private System.Windows.Forms.Label label2;
        private DevExpress.XtraEditors.LabelControl labelControl2;
        private System.Windows.Forms.ToolStripButton btnSungLuc;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator3;
		private System.Windows.Forms.ToolStripSeparator toolStripSeparator5;
		private System.Windows.Forms.ToolStripButton btnUpdateStepCode;
		private System.Windows.Forms.ToolStripButton btnSpeedReducer;
	}
}