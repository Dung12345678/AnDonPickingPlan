namespace BMS
{
    partial class frmGear
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
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmGear));
			this.colLotSize = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colID = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colWipGCQT = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colHYP = new DevExpress.XtraGrid.Columns.GridColumn();
			this.repositoryItemMemoEdit1 = new DevExpress.XtraEditors.Repository.RepositoryItemMemoEdit();
			this.colGear1 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.grvData = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.colGear2 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colGear3 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colLap = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colStockGCQT = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colGGiaCongThem = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colOnhandG = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colGThieu = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colWipG = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colGCQTThieu = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colWipTruc = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colOrderTrucMoi = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colHypNeed = new DevExpress.XtraGrid.Columns.GridColumn();
			this.grdData = new DevExpress.XtraGrid.GridControl();
			this.toolStripSeparator5 = new System.Windows.Forms.ToolStripSeparator();
			this.btnImportExcel = new System.Windows.Forms.ToolStripButton();
			this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
			this.btnDel = new System.Windows.Forms.ToolStripButton();
			this.toolStripSeparator4 = new System.Windows.Forms.ToolStripSeparator();
			this.btnEdit = new System.Windows.Forms.ToolStripButton();
			this.toolStripSeparator6 = new System.Windows.Forms.ToolStripSeparator();
			this.btnAdd = new System.Windows.Forms.ToolStripButton();
			this.toolStripSeparator7 = new System.Windows.Forms.ToolStripSeparator();
			this.btnCopy = new System.Windows.Forms.ToolStripButton();
			this.mnuMenu = new System.Windows.Forms.ToolStrip();
			this.btnExcel = new System.Windows.Forms.ToolStripButton();
			this.btnFindAll = new DevExpress.XtraEditors.SimpleButton();
			this.txtName = new System.Windows.Forms.TextBox();
			this.label2 = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemMemoEdit1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grvData)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grdData)).BeginInit();
			this.mnuMenu.SuspendLayout();
			this.SuspendLayout();
			// 
			// colLotSize
			// 
			this.colLotSize.AppearanceCell.Options.UseTextOptions = true;
			this.colLotSize.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colLotSize.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colLotSize.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colLotSize.AppearanceHeader.Options.UseBackColor = true;
			this.colLotSize.AppearanceHeader.Options.UseFont = true;
			this.colLotSize.AppearanceHeader.Options.UseTextOptions = true;
			this.colLotSize.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colLotSize.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colLotSize.Caption = "Lot size";
			this.colLotSize.FieldName = "LotSize";
			this.colLotSize.Name = "colLotSize";
			this.colLotSize.OptionsColumn.AllowEdit = false;
			this.colLotSize.Visible = true;
			this.colLotSize.VisibleIndex = 5;
			this.colLotSize.Width = 49;
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
			// colWipGCQT
			// 
			this.colWipGCQT.AppearanceCell.Options.UseTextOptions = true;
			this.colWipGCQT.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colWipGCQT.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colWipGCQT.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colWipGCQT.AppearanceHeader.Options.UseBackColor = true;
			this.colWipGCQT.AppearanceHeader.Options.UseFont = true;
			this.colWipGCQT.AppearanceHeader.Options.UseTextOptions = true;
			this.colWipGCQT.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colWipGCQT.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colWipGCQT.Caption = "WIP (GCQT)";
			this.colWipGCQT.FieldName = "WipGCQT";
			this.colWipGCQT.Name = "colWipGCQT";
			this.colWipGCQT.OptionsColumn.AllowEdit = false;
			this.colWipGCQT.Visible = true;
			this.colWipGCQT.VisibleIndex = 11;
			this.colWipGCQT.Width = 54;
			// 
			// colHYP
			// 
			this.colHYP.AppearanceCell.Options.UseTextOptions = true;
			this.colHYP.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colHYP.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colHYP.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colHYP.AppearanceHeader.Options.UseBackColor = true;
			this.colHYP.AppearanceHeader.Options.UseFont = true;
			this.colHYP.AppearanceHeader.Options.UseTextOptions = true;
			this.colHYP.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colHYP.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colHYP.Caption = "HYP";
			this.colHYP.FieldName = "HYP";
			this.colHYP.Name = "colHYP";
			this.colHYP.OptionsColumn.AllowEdit = false;
			this.colHYP.Summary.AddRange(new DevExpress.XtraGrid.GridSummaryItem[] {
            new DevExpress.XtraGrid.GridColumnSummaryItem(DevExpress.Data.SummaryItemType.Count)});
			this.colHYP.Visible = true;
			this.colHYP.VisibleIndex = 0;
			this.colHYP.Width = 118;
			// 
			// repositoryItemMemoEdit1
			// 
			this.repositoryItemMemoEdit1.Appearance.Options.UseTextOptions = true;
			this.repositoryItemMemoEdit1.Appearance.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.repositoryItemMemoEdit1.Name = "repositoryItemMemoEdit1";
			// 
			// colGear1
			// 
			this.colGear1.AppearanceCell.Options.UseTextOptions = true;
			this.colGear1.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colGear1.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colGear1.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colGear1.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colGear1.AppearanceHeader.Options.UseBackColor = true;
			this.colGear1.AppearanceHeader.Options.UseFont = true;
			this.colGear1.AppearanceHeader.Options.UseTextOptions = true;
			this.colGear1.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colGear1.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colGear1.Caption = "Gear 1";
			this.colGear1.FieldName = "Gear1";
			this.colGear1.Name = "colGear1";
			this.colGear1.OptionsColumn.AllowEdit = false;
			this.colGear1.Visible = true;
			this.colGear1.VisibleIndex = 1;
			this.colGear1.Width = 101;
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
            this.colHYP,
            this.colGear1,
            this.colGear2,
            this.colGear3,
            this.colLap,
            this.colWipGCQT,
            this.colID,
            this.colLotSize,
            this.colStockGCQT,
            this.colGGiaCongThem,
            this.colOnhandG,
            this.colGThieu,
            this.colWipG,
            this.colGCQTThieu,
            this.colWipTruc,
            this.colOrderTrucMoi,
            this.colHypNeed});
			this.grvData.CustomizationFormBounds = new System.Drawing.Rectangle(1293, 517, 210, 312);
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
			// colGear2
			// 
			this.colGear2.AppearanceCell.Options.UseTextOptions = true;
			this.colGear2.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colGear2.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colGear2.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colGear2.AppearanceHeader.Options.UseBackColor = true;
			this.colGear2.AppearanceHeader.Options.UseFont = true;
			this.colGear2.AppearanceHeader.Options.UseTextOptions = true;
			this.colGear2.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colGear2.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colGear2.Caption = "Gear 2";
			this.colGear2.FieldName = "Gear2";
			this.colGear2.Name = "colGear2";
			this.colGear2.OptionsColumn.AllowEdit = false;
			this.colGear2.Visible = true;
			this.colGear2.VisibleIndex = 2;
			this.colGear2.Width = 119;
			// 
			// colGear3
			// 
			this.colGear3.AppearanceCell.Options.UseTextOptions = true;
			this.colGear3.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colGear3.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colGear3.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colGear3.AppearanceHeader.Options.UseBackColor = true;
			this.colGear3.AppearanceHeader.Options.UseFont = true;
			this.colGear3.AppearanceHeader.Options.UseTextOptions = true;
			this.colGear3.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colGear3.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colGear3.Caption = "Gear 3";
			this.colGear3.FieldName = "Gear3";
			this.colGear3.Name = "colGear3";
			this.colGear3.OptionsColumn.AllowEdit = false;
			this.colGear3.Visible = true;
			this.colGear3.VisibleIndex = 3;
			this.colGear3.Width = 120;
			// 
			// colLap
			// 
			this.colLap.AppearanceCell.Options.UseTextOptions = true;
			this.colLap.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colLap.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colLap.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colLap.AppearanceHeader.Options.UseBackColor = true;
			this.colLap.AppearanceHeader.Options.UseFont = true;
			this.colLap.AppearanceHeader.Options.UseTextOptions = true;
			this.colLap.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colLap.Caption = "LAP";
			this.colLap.FieldName = "Lap";
			this.colLap.Name = "colLap";
			this.colLap.OptionsColumn.AllowEdit = false;
			this.colLap.Visible = true;
			this.colLap.VisibleIndex = 4;
			this.colLap.Width = 123;
			// 
			// colStockGCQT
			// 
			this.colStockGCQT.AppearanceCell.Options.UseTextOptions = true;
			this.colStockGCQT.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colStockGCQT.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colStockGCQT.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colStockGCQT.AppearanceHeader.Options.UseBackColor = true;
			this.colStockGCQT.AppearanceHeader.Options.UseFont = true;
			this.colStockGCQT.AppearanceHeader.Options.UseTextOptions = true;
			this.colStockGCQT.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colStockGCQT.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colStockGCQT.Caption = "Stock (GCQT)";
			this.colStockGCQT.FieldName = "StockGCQT";
			this.colStockGCQT.Name = "colStockGCQT";
			this.colStockGCQT.OptionsColumn.AllowEdit = false;
			this.colStockGCQT.Visible = true;
			this.colStockGCQT.VisibleIndex = 10;
			this.colStockGCQT.Width = 56;
			// 
			// colGGiaCongThem
			// 
			this.colGGiaCongThem.AppearanceCell.Options.UseTextOptions = true;
			this.colGGiaCongThem.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colGGiaCongThem.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colGGiaCongThem.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colGGiaCongThem.AppearanceHeader.Options.UseBackColor = true;
			this.colGGiaCongThem.AppearanceHeader.Options.UseFont = true;
			this.colGGiaCongThem.AppearanceHeader.Options.UseTextOptions = true;
			this.colGGiaCongThem.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colGGiaCongThem.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colGGiaCongThem.Caption = "(G) Gia công thêm";
			this.colGGiaCongThem.FieldName = "GGiaCongThem";
			this.colGGiaCongThem.Name = "colGGiaCongThem";
			this.colGGiaCongThem.OptionsColumn.AllowEdit = false;
			this.colGGiaCongThem.Visible = true;
			this.colGGiaCongThem.VisibleIndex = 7;
			// 
			// colOnhandG
			// 
			this.colOnhandG.AppearanceCell.Options.UseTextOptions = true;
			this.colOnhandG.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colOnhandG.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colOnhandG.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colOnhandG.AppearanceHeader.Options.UseBackColor = true;
			this.colOnhandG.AppearanceHeader.Options.UseFont = true;
			this.colOnhandG.AppearanceHeader.Options.UseTextOptions = true;
			this.colOnhandG.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colOnhandG.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colOnhandG.Caption = "ON HAND G";
			this.colOnhandG.FieldName = "OnhandG";
			this.colOnhandG.Name = "colOnhandG";
			this.colOnhandG.OptionsColumn.AllowEdit = false;
			this.colOnhandG.Visible = true;
			this.colOnhandG.VisibleIndex = 9;
			this.colOnhandG.Width = 49;
			// 
			// colGThieu
			// 
			this.colGThieu.AppearanceCell.Options.UseTextOptions = true;
			this.colGThieu.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colGThieu.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colGThieu.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colGThieu.AppearanceHeader.Options.UseBackColor = true;
			this.colGThieu.AppearanceHeader.Options.UseFont = true;
			this.colGThieu.AppearanceHeader.Options.UseTextOptions = true;
			this.colGThieu.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colGThieu.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colGThieu.Caption = "(G) Thieu";
			this.colGThieu.FieldName = "GThieu";
			this.colGThieu.Name = "colGThieu";
			this.colGThieu.OptionsColumn.AllowEdit = false;
			this.colGThieu.Visible = true;
			this.colGThieu.VisibleIndex = 6;
			this.colGThieu.Width = 57;
			// 
			// colWipG
			// 
			this.colWipG.AppearanceCell.Options.UseTextOptions = true;
			this.colWipG.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colWipG.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colWipG.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colWipG.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colWipG.AppearanceHeader.Options.UseBackColor = true;
			this.colWipG.AppearanceHeader.Options.UseFont = true;
			this.colWipG.AppearanceHeader.Options.UseTextOptions = true;
			this.colWipG.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colWipG.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colWipG.Caption = "WIP (G)";
			this.colWipG.ColumnEdit = this.repositoryItemMemoEdit1;
			this.colWipG.FieldName = "WipG";
			this.colWipG.Name = "colWipG";
			this.colWipG.OptionsColumn.AllowEdit = false;
			this.colWipG.Visible = true;
			this.colWipG.VisibleIndex = 8;
			this.colWipG.Width = 49;
			// 
			// colGCQTThieu
			// 
			this.colGCQTThieu.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colGCQTThieu.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colGCQTThieu.AppearanceHeader.Options.UseBackColor = true;
			this.colGCQTThieu.AppearanceHeader.Options.UseFont = true;
			this.colGCQTThieu.AppearanceHeader.Options.UseTextOptions = true;
			this.colGCQTThieu.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colGCQTThieu.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colGCQTThieu.Caption = "(GCQT) thieu";
			this.colGCQTThieu.FieldName = "GCQTThieu";
			this.colGCQTThieu.Name = "colGCQTThieu";
			this.colGCQTThieu.OptionsColumn.AllowEdit = false;
			this.colGCQTThieu.Visible = true;
			this.colGCQTThieu.VisibleIndex = 12;
			this.colGCQTThieu.Width = 52;
			// 
			// colWipTruc
			// 
			this.colWipTruc.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colWipTruc.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colWipTruc.AppearanceHeader.Options.UseBackColor = true;
			this.colWipTruc.AppearanceHeader.Options.UseFont = true;
			this.colWipTruc.AppearanceHeader.Options.UseTextOptions = true;
			this.colWipTruc.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colWipTruc.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colWipTruc.Caption = "WIP truc";
			this.colWipTruc.FieldName = "WipTruc";
			this.colWipTruc.Name = "colWipTruc";
			this.colWipTruc.OptionsColumn.AllowEdit = false;
			this.colWipTruc.Visible = true;
			this.colWipTruc.VisibleIndex = 13;
			this.colWipTruc.Width = 48;
			// 
			// colOrderTrucMoi
			// 
			this.colOrderTrucMoi.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colOrderTrucMoi.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colOrderTrucMoi.AppearanceHeader.Options.UseBackColor = true;
			this.colOrderTrucMoi.AppearanceHeader.Options.UseFont = true;
			this.colOrderTrucMoi.AppearanceHeader.Options.UseTextOptions = true;
			this.colOrderTrucMoi.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colOrderTrucMoi.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colOrderTrucMoi.Caption = "Order truc moi";
			this.colOrderTrucMoi.FieldName = "OrderTrucMoi";
			this.colOrderTrucMoi.Name = "colOrderTrucMoi";
			this.colOrderTrucMoi.OptionsColumn.AllowEdit = false;
			this.colOrderTrucMoi.Visible = true;
			this.colOrderTrucMoi.VisibleIndex = 14;
			this.colOrderTrucMoi.Width = 57;
			// 
			// colHypNeed
			// 
			this.colHypNeed.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colHypNeed.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colHypNeed.AppearanceHeader.Options.UseBackColor = true;
			this.colHypNeed.AppearanceHeader.Options.UseFont = true;
			this.colHypNeed.AppearanceHeader.Options.UseTextOptions = true;
			this.colHypNeed.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colHypNeed.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colHypNeed.Caption = "HYP cần";
			this.colHypNeed.FieldName = "HypNeed";
			this.colHypNeed.Name = "colHypNeed";
			this.colHypNeed.OptionsColumn.AllowEdit = false;
			this.colHypNeed.Visible = true;
			this.colHypNeed.VisibleIndex = 15;
			this.colHypNeed.Width = 53;
			// 
			// grdData
			// 
			this.grdData.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.grdData.Font = new System.Drawing.Font("Times New Roman", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.grdData.Location = new System.Drawing.Point(0, 74);
			this.grdData.MainView = this.grvData;
			this.grdData.Name = "grdData";
			this.grdData.RepositoryItems.AddRange(new DevExpress.XtraEditors.Repository.RepositoryItem[] {
            this.repositoryItemMemoEdit1});
			this.grdData.Size = new System.Drawing.Size(1206, 602);
			this.grdData.TabIndex = 16;
			this.grdData.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.grvData});
			// 
			// toolStripSeparator5
			// 
			this.toolStripSeparator5.AutoSize = false;
			this.toolStripSeparator5.Name = "toolStripSeparator5";
			this.toolStripSeparator5.Size = new System.Drawing.Size(6, 40);
			// 
			// btnImportExcel
			// 
			this.btnImportExcel.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnImportExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnImportExcel.Image")));
			this.btnImportExcel.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnImportExcel.Name = "btnImportExcel";
			this.btnImportExcel.Size = new System.Drawing.Size(71, 33);
			this.btnImportExcel.Tag = "";
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
			this.btnDel.Tag = "";
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
			this.btnEdit.Tag = "";
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
			this.btnAdd.Tag = "";
			this.btnAdd.Text = "Tạo sản phẩm";
			this.btnAdd.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
			// 
			// toolStripSeparator7
			// 
			this.toolStripSeparator7.Name = "toolStripSeparator7";
			this.toolStripSeparator7.Size = new System.Drawing.Size(6, 23);
			this.toolStripSeparator7.Visible = false;
			// 
			// btnCopy
			// 
			this.btnCopy.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnCopy.Image = ((System.Drawing.Image)(resources.GetObject("btnCopy.Image")));
			this.btnCopy.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnCopy.Name = "btnCopy";
			this.btnCopy.Size = new System.Drawing.Size(39, 33);
			this.btnCopy.Tag = "";
			this.btnCopy.Text = "Copy";
			this.btnCopy.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnCopy.Visible = false;
			this.btnCopy.Click += new System.EventHandler(this.btnCopy_Click);
			// 
			// mnuMenu
			// 
			this.mnuMenu.AutoSize = false;
			this.mnuMenu.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(234)))), ((int)(((byte)(234)))), ((int)(((byte)(234)))));
			this.mnuMenu.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.mnuMenu.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden;
			this.mnuMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.btnAdd,
            this.toolStripSeparator6,
            this.btnEdit,
            this.toolStripSeparator4,
            this.btnDel,
            this.toolStripSeparator2,
            this.btnExcel,
            this.toolStripSeparator5,
            this.btnImportExcel,
            this.btnCopy,
            this.toolStripSeparator7});
			this.mnuMenu.LayoutStyle = System.Windows.Forms.ToolStripLayoutStyle.Flow;
			this.mnuMenu.Location = new System.Drawing.Point(0, 0);
			this.mnuMenu.Name = "mnuMenu";
			this.mnuMenu.RenderMode = System.Windows.Forms.ToolStripRenderMode.Professional;
			this.mnuMenu.Size = new System.Drawing.Size(1206, 42);
			this.mnuMenu.TabIndex = 20;
			this.mnuMenu.Text = "toolStrip2";
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
			// btnFindAll
			// 
			this.btnFindAll.Location = new System.Drawing.Point(372, 46);
			this.btnFindAll.Name = "btnFindAll";
			this.btnFindAll.Size = new System.Drawing.Size(75, 23);
			this.btnFindAll.TabIndex = 21;
			this.btnFindAll.Text = "Tìm kiếm";
			this.btnFindAll.Click += new System.EventHandler(this.btnFindAll_Click);
			// 
			// txtName
			// 
			this.txtName.Location = new System.Drawing.Point(77, 48);
			this.txtName.Name = "txtName";
			this.txtName.Size = new System.Drawing.Size(290, 20);
			this.txtName.TabIndex = 20;
			this.txtName.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtName_KeyDown);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(6, 51);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(66, 13);
			this.label2.TabIndex = 19;
			this.label2.Text = "Tên module:";
			// 
			// frmGear
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(1206, 677);
			this.Controls.Add(this.btnFindAll);
			this.Controls.Add(this.txtName);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.mnuMenu);
			this.Controls.Add(this.grdData);
			this.Name = "frmGear";
			this.Text = "DANH SÁCH SẢN PHẨM";
			this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
			this.Load += new System.EventHandler(this.frmProduct_Load);
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemMemoEdit1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grvData)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grdData)).EndInit();
			this.mnuMenu.ResumeLayout(false);
			this.mnuMenu.PerformLayout();
			this.ResumeLayout(false);
			this.PerformLayout();

        }

        #endregion

        private DevExpress.XtraGrid.Columns.GridColumn colLotSize;
        private DevExpress.XtraGrid.Columns.GridColumn colID;
        private DevExpress.XtraGrid.Columns.GridColumn colWipGCQT;
        private DevExpress.XtraGrid.Columns.GridColumn colHYP;
        private DevExpress.XtraEditors.Repository.RepositoryItemMemoEdit repositoryItemMemoEdit1;
        private DevExpress.XtraGrid.Columns.GridColumn colGear1;
        private DevExpress.XtraGrid.Views.Grid.GridView grvData;
        private DevExpress.XtraGrid.Columns.GridColumn colStockGCQT;
        private DevExpress.XtraGrid.Columns.GridColumn colGGiaCongThem;
        private DevExpress.XtraGrid.Columns.GridColumn colOnhandG;
        private DevExpress.XtraGrid.Columns.GridColumn colGear2;
        private DevExpress.XtraGrid.Columns.GridColumn colGear3;
        private DevExpress.XtraGrid.Columns.GridColumn colLap;
        private DevExpress.XtraGrid.Columns.GridColumn colGThieu;
        private DevExpress.XtraGrid.Columns.GridColumn colWipG;
        private DevExpress.XtraGrid.GridControl grdData;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator5;
        private System.Windows.Forms.ToolStripButton btnImportExcel;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
        private System.Windows.Forms.ToolStripButton btnDel;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator4;
        private System.Windows.Forms.ToolStripButton btnEdit;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator6;
        private System.Windows.Forms.ToolStripButton btnAdd;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator7;
        private System.Windows.Forms.ToolStripButton btnCopy;
        private System.Windows.Forms.ToolStrip mnuMenu;
        private System.Windows.Forms.ToolStripButton btnExcel;
        private DevExpress.XtraEditors.SimpleButton btnFindAll;
        private System.Windows.Forms.TextBox txtName;
        private System.Windows.Forms.Label label2;
        private DevExpress.XtraGrid.Columns.GridColumn colGCQTThieu;
        private DevExpress.XtraGrid.Columns.GridColumn colWipTruc;
        private DevExpress.XtraGrid.Columns.GridColumn colOrderTrucMoi;
        private DevExpress.XtraGrid.Columns.GridColumn colHypNeed;
    }
}