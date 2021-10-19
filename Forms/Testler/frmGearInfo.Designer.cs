namespace BMS
{
    partial class frmGearInfo
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
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmGearInfo));
			this.colSlitMin = new DevExpress.XtraGrid.Columns.GridColumn();
			this.repositoryItemTextEdit1 = new DevExpress.XtraEditors.Repository.RepositoryItemTextEdit();
			this.colID = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colWipGCQT = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colGearCode = new DevExpress.XtraGrid.Columns.GridColumn();
			this.repositoryItemMemoEdit1 = new DevExpress.XtraEditors.Repository.RepositoryItemMemoEdit();
			this.grvData = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.colStockGCQT = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colVibrateMin = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colOnhandG = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colSlitMax = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colVibrateMax = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colGCQTThieu = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colWipTruc = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colOrderTrucMoi = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colStatus = new DevExpress.XtraGrid.Columns.GridColumn();
			this.grdData = new DevExpress.XtraGrid.GridControl();
			this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
			this.btnDel = new System.Windows.Forms.ToolStripButton();
			this.btnReload = new System.Windows.Forms.ToolStripButton();
			this.mnuMenu = new System.Windows.Forms.ToolStrip();
			this.btnSave = new System.Windows.Forms.ToolStripButton();
			this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemTextEdit1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemMemoEdit1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grvData)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grdData)).BeginInit();
			this.mnuMenu.SuspendLayout();
			this.SuspendLayout();
			// 
			// colSlitMin
			// 
			this.colSlitMin.AppearanceCell.Options.UseTextOptions = true;
			this.colSlitMin.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colSlitMin.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colSlitMin.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colSlitMin.AppearanceHeader.Options.UseBackColor = true;
			this.colSlitMin.AppearanceHeader.Options.UseFont = true;
			this.colSlitMin.AppearanceHeader.Options.UseTextOptions = true;
			this.colSlitMin.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colSlitMin.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colSlitMin.Caption = "Khe hở MIN";
			this.colSlitMin.ColumnEdit = this.repositoryItemTextEdit1;
			this.colSlitMin.FieldName = "SlitMin";
			this.colSlitMin.Name = "colSlitMin";
			this.colSlitMin.Visible = true;
			this.colSlitMin.VisibleIndex = 1;
			this.colSlitMin.Width = 107;
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
			// 
			// colGearCode
			// 
			this.colGearCode.AppearanceCell.Options.UseTextOptions = true;
			this.colGearCode.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colGearCode.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colGearCode.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colGearCode.AppearanceHeader.Options.UseBackColor = true;
			this.colGearCode.AppearanceHeader.Options.UseFont = true;
			this.colGearCode.AppearanceHeader.Options.UseTextOptions = true;
			this.colGearCode.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colGearCode.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colGearCode.Caption = "Mã hàng";
			this.colGearCode.FieldName = "GearCode";
			this.colGearCode.Name = "colGearCode";
			this.colGearCode.Summary.AddRange(new DevExpress.XtraGrid.GridSummaryItem[] {
            new DevExpress.XtraGrid.GridColumnSummaryItem(DevExpress.Data.SummaryItemType.Count)});
			this.colGearCode.Visible = true;
			this.colGearCode.VisibleIndex = 0;
			this.colGearCode.Width = 200;
			// 
			// repositoryItemMemoEdit1
			// 
			this.repositoryItemMemoEdit1.Appearance.Options.UseTextOptions = true;
			this.repositoryItemMemoEdit1.Appearance.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.repositoryItemMemoEdit1.Name = "repositoryItemMemoEdit1";
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
            this.colID,
            this.colGearCode,
            this.colWipGCQT,
            this.colSlitMin,
            this.colStockGCQT,
            this.colVibrateMin,
            this.colOnhandG,
            this.colSlitMax,
            this.colVibrateMax,
            this.colGCQTThieu,
            this.colWipTruc,
            this.colOrderTrucMoi,
            this.colStatus});
			this.grvData.CustomizationFormBounds = new System.Drawing.Rectangle(1293, 517, 210, 312);
			this.grvData.GridControl = this.grdData;
			this.grvData.HorzScrollStep = 5;
			this.grvData.Name = "grvData";
			this.grvData.OptionsBehavior.AutoExpandAllGroups = true;
			this.grvData.OptionsCustomization.AllowRowSizing = true;
			this.grvData.OptionsFind.AlwaysVisible = true;
			this.grvData.OptionsFind.ShowCloseButton = false;
			this.grvData.OptionsView.ColumnAutoWidth = false;
			this.grvData.OptionsView.NewItemRowPosition = DevExpress.XtraGrid.Views.Grid.NewItemRowPosition.Top;
			this.grvData.OptionsView.RowAutoHeight = true;
			this.grvData.OptionsView.ShowFooter = true;
			this.grvData.OptionsView.ShowGroupPanel = false;
			this.grvData.CellValueChanged += new DevExpress.XtraGrid.Views.Base.CellValueChangedEventHandler(this.grvData_CellValueChanged);
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
			// 
			// colVibrateMin
			// 
			this.colVibrateMin.AppearanceCell.Options.UseTextOptions = true;
			this.colVibrateMin.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colVibrateMin.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colVibrateMin.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colVibrateMin.AppearanceHeader.Options.UseBackColor = true;
			this.colVibrateMin.AppearanceHeader.Options.UseFont = true;
			this.colVibrateMin.AppearanceHeader.Options.UseTextOptions = true;
			this.colVibrateMin.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colVibrateMin.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colVibrateMin.Caption = "Rung động MIN";
			this.colVibrateMin.ColumnEdit = this.repositoryItemTextEdit1;
			this.colVibrateMin.FieldName = "VibrateMin";
			this.colVibrateMin.Name = "colVibrateMin";
			this.colVibrateMin.Visible = true;
			this.colVibrateMin.VisibleIndex = 3;
			this.colVibrateMin.Width = 128;
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
			this.colOnhandG.Width = 76;
			// 
			// colSlitMax
			// 
			this.colSlitMax.AppearanceCell.Options.UseTextOptions = true;
			this.colSlitMax.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colSlitMax.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colSlitMax.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colSlitMax.AppearanceHeader.Options.UseBackColor = true;
			this.colSlitMax.AppearanceHeader.Options.UseFont = true;
			this.colSlitMax.AppearanceHeader.Options.UseTextOptions = true;
			this.colSlitMax.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colSlitMax.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colSlitMax.Caption = "Khe hở MAX";
			this.colSlitMax.ColumnEdit = this.repositoryItemTextEdit1;
			this.colSlitMax.FieldName = "SlitMax";
			this.colSlitMax.Name = "colSlitMax";
			this.colSlitMax.Visible = true;
			this.colSlitMax.VisibleIndex = 2;
			this.colSlitMax.Width = 117;
			// 
			// colVibrateMax
			// 
			this.colVibrateMax.AppearanceCell.Options.UseTextOptions = true;
			this.colVibrateMax.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colVibrateMax.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colVibrateMax.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colVibrateMax.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colVibrateMax.AppearanceHeader.Options.UseBackColor = true;
			this.colVibrateMax.AppearanceHeader.Options.UseFont = true;
			this.colVibrateMax.AppearanceHeader.Options.UseTextOptions = true;
			this.colVibrateMax.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colVibrateMax.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colVibrateMax.Caption = "Rung động MAX";
			this.colVibrateMax.ColumnEdit = this.repositoryItemTextEdit1;
			this.colVibrateMax.FieldName = "VibrateMax";
			this.colVibrateMax.Name = "colVibrateMax";
			this.colVibrateMax.Visible = true;
			this.colVibrateMax.VisibleIndex = 4;
			this.colVibrateMax.Width = 117;
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
			// 
			// colStatus
			// 
			this.colStatus.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colStatus.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colStatus.AppearanceHeader.Options.UseBackColor = true;
			this.colStatus.AppearanceHeader.Options.UseFont = true;
			this.colStatus.AppearanceHeader.Options.UseTextOptions = true;
			this.colStatus.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colStatus.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colStatus.Caption = "Status";
			this.colStatus.FieldName = "Status";
			this.colStatus.Name = "colStatus";
			this.colStatus.OptionsColumn.AllowEdit = false;
			// 
			// grdData
			// 
			this.grdData.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.grdData.Font = new System.Drawing.Font("Times New Roman", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.grdData.Location = new System.Drawing.Point(0, 45);
			this.grdData.MainView = this.grvData;
			this.grdData.Name = "grdData";
			this.grdData.RepositoryItems.AddRange(new DevExpress.XtraEditors.Repository.RepositoryItem[] {
            this.repositoryItemMemoEdit1,
            this.repositoryItemTextEdit1});
			this.grdData.Size = new System.Drawing.Size(758, 695);
			this.grdData.TabIndex = 16;
			this.grdData.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.grvData});
			// 
			// toolStripSeparator2
			// 
			this.toolStripSeparator2.AutoSize = false;
			this.toolStripSeparator2.Name = "toolStripSeparator2";
			this.toolStripSeparator2.Size = new System.Drawing.Size(6, 40);
			// 
			// btnDel
			// 
			this.btnDel.AutoSize = false;
			this.btnDel.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnDel.Image = ((System.Drawing.Image)(resources.GetObject("btnDel.Image")));
			this.btnDel.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnDel.Name = "btnDel";
			this.btnDel.Size = new System.Drawing.Size(80, 33);
			this.btnDel.Tag = "";
			this.btnDel.Text = "Xóa";
			this.btnDel.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnDel.Click += new System.EventHandler(this.btnDel_Click);
			// 
			// btnReload
			// 
			this.btnReload.AutoSize = false;
			this.btnReload.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnReload.Image = ((System.Drawing.Image)(resources.GetObject("btnReload.Image")));
			this.btnReload.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnReload.Name = "btnReload";
			this.btnReload.Size = new System.Drawing.Size(70, 33);
			this.btnReload.Tag = "";
			this.btnReload.Text = "Reload";
			this.btnReload.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnReload.Click += new System.EventHandler(this.btnReload_Click);
			// 
			// mnuMenu
			// 
			this.mnuMenu.AutoSize = false;
			this.mnuMenu.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(234)))), ((int)(((byte)(234)))), ((int)(((byte)(234)))));
			this.mnuMenu.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.mnuMenu.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden;
			this.mnuMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.btnDel,
            this.toolStripSeparator2,
            this.btnSave,
            this.toolStripSeparator1,
            this.btnReload});
			this.mnuMenu.LayoutStyle = System.Windows.Forms.ToolStripLayoutStyle.Flow;
			this.mnuMenu.Location = new System.Drawing.Point(0, 0);
			this.mnuMenu.Name = "mnuMenu";
			this.mnuMenu.RenderMode = System.Windows.Forms.ToolStripRenderMode.Professional;
			this.mnuMenu.Size = new System.Drawing.Size(758, 42);
			this.mnuMenu.TabIndex = 20;
			this.mnuMenu.Text = "toolStrip2";
			// 
			// btnSave
			// 
			this.btnSave.AutoSize = false;
			this.btnSave.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnSave.Image = ((System.Drawing.Image)(resources.GetObject("btnSave.Image")));
			this.btnSave.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnSave.Name = "btnSave";
			this.btnSave.Size = new System.Drawing.Size(70, 33);
			this.btnSave.Text = "Cất";
			this.btnSave.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
			// 
			// toolStripSeparator1
			// 
			this.toolStripSeparator1.AutoSize = false;
			this.toolStripSeparator1.Name = "toolStripSeparator1";
			this.toolStripSeparator1.Size = new System.Drawing.Size(6, 40);
			// 
			// frmGearInfo
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(758, 741);
			this.Controls.Add(this.mnuMenu);
			this.Controls.Add(this.grdData);
			this.Name = "frmGearInfo";
			this.Text = "KHE HỞ -RUNG";
			this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
			this.Load += new System.EventHandler(this.frmProduct_Load);
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemTextEdit1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemMemoEdit1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grvData)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grdData)).EndInit();
			this.mnuMenu.ResumeLayout(false);
			this.mnuMenu.PerformLayout();
			this.ResumeLayout(false);

        }

        #endregion

        private DevExpress.XtraGrid.Columns.GridColumn colSlitMin;
        private DevExpress.XtraGrid.Columns.GridColumn colID;
        private DevExpress.XtraGrid.Columns.GridColumn colWipGCQT;
        private DevExpress.XtraGrid.Columns.GridColumn colGearCode;
        private DevExpress.XtraEditors.Repository.RepositoryItemMemoEdit repositoryItemMemoEdit1;
        private DevExpress.XtraGrid.Views.Grid.GridView grvData;
        private DevExpress.XtraGrid.Columns.GridColumn colStockGCQT;
        private DevExpress.XtraGrid.Columns.GridColumn colVibrateMin;
        private DevExpress.XtraGrid.Columns.GridColumn colOnhandG;
        private DevExpress.XtraGrid.Columns.GridColumn colSlitMax;
        private DevExpress.XtraGrid.Columns.GridColumn colVibrateMax;
        private DevExpress.XtraGrid.GridControl grdData;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
        private System.Windows.Forms.ToolStripButton btnDel;
        private System.Windows.Forms.ToolStripButton btnReload;
        private System.Windows.Forms.ToolStrip mnuMenu;
        private DevExpress.XtraGrid.Columns.GridColumn colGCQTThieu;
        private DevExpress.XtraGrid.Columns.GridColumn colWipTruc;
        private DevExpress.XtraGrid.Columns.GridColumn colOrderTrucMoi;
        private DevExpress.XtraGrid.Columns.GridColumn colStatus;
        private System.Windows.Forms.ToolStripButton btnSave;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private DevExpress.XtraEditors.Repository.RepositoryItemTextEdit repositoryItemTextEdit1;
    }
}