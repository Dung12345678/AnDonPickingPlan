namespace BMS
{
    partial class frmHistoryPartOutDetail
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
			this.gridColumn1 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colValueTypeText = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colRealValue1 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.grdData = new DevExpress.XtraGrid.GridControl();
			this.grvData = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.colID = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colStockID = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colShelf = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colLo = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colArticleID = new DevExpress.XtraGrid.Columns.GridColumn();
			this.repositoryItemMemoEdit1 = new DevExpress.XtraEditors.Repository.RepositoryItemMemoEdit();
			this.colAccessory = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colDescription = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colCreatDate = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colQty = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colRealQty = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colStockCDID = new DevExpress.XtraGrid.Columns.GridColumn();
			this.col1 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.col2 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.col3 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.col4 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.col5 = new DevExpress.XtraGrid.Columns.GridColumn();
			((System.ComponentModel.ISupportInitialize)(this.grdData)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grvData)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemMemoEdit1)).BeginInit();
			this.SuspendLayout();
			// 
			// gridColumn1
			// 
			this.gridColumn1.AppearanceCell.Options.UseTextOptions = true;
			this.gridColumn1.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.gridColumn1.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.gridColumn1.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.gridColumn1.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.gridColumn1.AppearanceHeader.Options.UseFont = true;
			this.gridColumn1.AppearanceHeader.Options.UseTextOptions = true;
			this.gridColumn1.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.gridColumn1.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.gridColumn1.Caption = "Giá trị tiêu chuẩn";
			this.gridColumn1.FieldName = "StandardValue";
			this.gridColumn1.Name = "gridColumn1";
			this.gridColumn1.OptionsColumn.AllowEdit = false;
			this.gridColumn1.OptionsColumn.AllowFocus = false;
			this.gridColumn1.Visible = true;
			this.gridColumn1.VisibleIndex = 1;
			this.gridColumn1.Width = 100;
			// 
			// colValueTypeText
			// 
			this.colValueTypeText.AppearanceCell.Options.UseTextOptions = true;
			this.colValueTypeText.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colValueTypeText.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colValueTypeText.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colValueTypeText.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colValueTypeText.AppearanceHeader.Options.UseFont = true;
			this.colValueTypeText.Caption = "Ghi chép";
			this.colValueTypeText.FieldName = "ValueTypeText";
			this.colValueTypeText.Name = "colValueTypeText";
			this.colValueTypeText.OptionsColumn.AllowEdit = false;
			this.colValueTypeText.OptionsColumn.AllowFocus = false;
			this.colValueTypeText.Visible = true;
			this.colValueTypeText.VisibleIndex = 2;
			this.colValueTypeText.Width = 70;
			// 
			// colRealValue1
			// 
			this.colRealValue1.AppearanceCell.Options.UseTextOptions = true;
			this.colRealValue1.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colRealValue1.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colRealValue1.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colRealValue1.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Bold);
			this.colRealValue1.AppearanceHeader.Options.UseFont = true;
			this.colRealValue1.AppearanceHeader.Options.UseTextOptions = true;
			this.colRealValue1.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colRealValue1.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colRealValue1.Caption = "#";
			this.colRealValue1.FieldName = "RealValue1";
			this.colRealValue1.Name = "colRealValue1";
			this.colRealValue1.Visible = true;
			this.colRealValue1.VisibleIndex = 3;
			this.colRealValue1.Width = 85;
			// 
			// grdData
			// 
			this.grdData.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.grdData.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.grdData.Location = new System.Drawing.Point(3, 0);
			this.grdData.MainView = this.grvData;
			this.grdData.Name = "grdData";
			this.grdData.RepositoryItems.AddRange(new DevExpress.XtraEditors.Repository.RepositoryItem[] {
            this.repositoryItemMemoEdit1});
			this.grdData.Size = new System.Drawing.Size(1018, 533);
			this.grdData.TabIndex = 30;
			this.grdData.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.grvData});
			// 
			// grvData
			// 
			this.grvData.ColumnPanelRowHeight = 40;
			this.grvData.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
            this.colID,
            this.colStockID,
            this.colShelf,
            this.colLo,
            this.colArticleID,
            this.colAccessory,
            this.colDescription,
            this.colCreatDate,
            this.colQty,
            this.colRealQty,
            this.colStockCDID,
            this.col1,
            this.col2,
            this.col3,
            this.col4,
            this.col5});
			this.grvData.GridControl = this.grdData;
			this.grvData.Name = "grvData";
			this.grvData.OptionsBehavior.AutoExpandAllGroups = true;
			this.grvData.OptionsBehavior.Editable = false;
			this.grvData.OptionsSelection.EnableAppearanceFocusedCell = false;
			this.grvData.OptionsSelection.EnableAppearanceFocusedRow = false;
			this.grvData.OptionsView.ColumnAutoWidth = false;
			this.grvData.OptionsView.RowAutoHeight = true;
			this.grvData.OptionsView.ShowGroupPanel = false;
			this.grvData.PaintStyleName = "Flat";
			this.grvData.RowHeight = 25;
			// 
			// colID
			// 
			this.colID.AppearanceCell.Options.UseTextOptions = true;
			this.colID.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colID.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
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
			this.colStockID.Caption = "StockID";
			this.colStockID.FieldName = "StockID";
			this.colStockID.Name = "colStockID";
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
			this.colShelf.AppearanceHeader.BackColor = System.Drawing.Color.DodgerBlue;
			this.colShelf.AppearanceHeader.BackColor2 = System.Drawing.Color.DodgerBlue;
			this.colShelf.AppearanceHeader.BorderColor = System.Drawing.Color.DodgerBlue;
			this.colShelf.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colShelf.AppearanceHeader.ForeColor = System.Drawing.Color.White;
			this.colShelf.AppearanceHeader.Options.UseBackColor = true;
			this.colShelf.AppearanceHeader.Options.UseBorderColor = true;
			this.colShelf.AppearanceHeader.Options.UseFont = true;
			this.colShelf.AppearanceHeader.Options.UseForeColor = true;
			this.colShelf.AppearanceHeader.Options.UseTextOptions = true;
			this.colShelf.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colShelf.Caption = "Vị trí";
			this.colShelf.FieldName = "Shelf";
			this.colShelf.Name = "colShelf";
			this.colShelf.OptionsColumn.AllowEdit = false;
			this.colShelf.Visible = true;
			this.colShelf.VisibleIndex = 2;
			this.colShelf.Width = 89;
			// 
			// colLo
			// 
			this.colLo.AppearanceCell.BackColor = System.Drawing.Color.White;
			this.colLo.AppearanceCell.BackColor2 = System.Drawing.Color.White;
			this.colLo.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.colLo.AppearanceCell.Options.UseBackColor = true;
			this.colLo.AppearanceCell.Options.UseFont = true;
			this.colLo.AppearanceCell.Options.UseTextOptions = true;
			this.colLo.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colLo.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colLo.AppearanceHeader.BackColor = System.Drawing.Color.DodgerBlue;
			this.colLo.AppearanceHeader.BackColor2 = System.Drawing.Color.DodgerBlue;
			this.colLo.AppearanceHeader.BorderColor = System.Drawing.Color.DodgerBlue;
			this.colLo.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colLo.AppearanceHeader.ForeColor = System.Drawing.Color.White;
			this.colLo.AppearanceHeader.Options.UseBackColor = true;
			this.colLo.AppearanceHeader.Options.UseBorderColor = true;
			this.colLo.AppearanceHeader.Options.UseFont = true;
			this.colLo.AppearanceHeader.Options.UseForeColor = true;
			this.colLo.AppearanceHeader.Options.UseTextOptions = true;
			this.colLo.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colLo.Caption = "Lo";
			this.colLo.FieldName = "Lo";
			this.colLo.Name = "colLo";
			this.colLo.OptionsColumn.AllowEdit = false;
			this.colLo.Visible = true;
			this.colLo.VisibleIndex = 4;
			this.colLo.Width = 127;
			// 
			// colArticleID
			// 
			this.colArticleID.AppearanceCell.BackColor = System.Drawing.Color.White;
			this.colArticleID.AppearanceCell.BackColor2 = System.Drawing.Color.White;
			this.colArticleID.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.colArticleID.AppearanceCell.Options.UseBackColor = true;
			this.colArticleID.AppearanceCell.Options.UseFont = true;
			this.colArticleID.AppearanceCell.Options.UseTextOptions = true;
			this.colArticleID.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colArticleID.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colArticleID.AppearanceHeader.BackColor = System.Drawing.Color.DodgerBlue;
			this.colArticleID.AppearanceHeader.BackColor2 = System.Drawing.Color.DodgerBlue;
			this.colArticleID.AppearanceHeader.BorderColor = System.Drawing.Color.DodgerBlue;
			this.colArticleID.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colArticleID.AppearanceHeader.ForeColor = System.Drawing.Color.White;
			this.colArticleID.AppearanceHeader.Options.UseBackColor = true;
			this.colArticleID.AppearanceHeader.Options.UseBorderColor = true;
			this.colArticleID.AppearanceHeader.Options.UseFont = true;
			this.colArticleID.AppearanceHeader.Options.UseForeColor = true;
			this.colArticleID.AppearanceHeader.Options.UseTextOptions = true;
			this.colArticleID.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colArticleID.Caption = "ArticleID";
			this.colArticleID.ColumnEdit = this.repositoryItemMemoEdit1;
			this.colArticleID.FieldName = "ArticleID";
			this.colArticleID.Name = "colArticleID";
			this.colArticleID.OptionsColumn.AllowEdit = false;
			this.colArticleID.Visible = true;
			this.colArticleID.VisibleIndex = 0;
			this.colArticleID.Width = 108;
			// 
			// repositoryItemMemoEdit1
			// 
			this.repositoryItemMemoEdit1.Name = "repositoryItemMemoEdit1";
			// 
			// colAccessory
			// 
			this.colAccessory.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.colAccessory.AppearanceCell.Options.UseFont = true;
			this.colAccessory.AppearanceCell.Options.UseTextOptions = true;
			this.colAccessory.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colAccessory.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colAccessory.AppearanceHeader.BackColor = System.Drawing.Color.DodgerBlue;
			this.colAccessory.AppearanceHeader.BackColor2 = System.Drawing.Color.DodgerBlue;
			this.colAccessory.AppearanceHeader.BorderColor = System.Drawing.Color.DodgerBlue;
			this.colAccessory.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colAccessory.AppearanceHeader.ForeColor = System.Drawing.Color.White;
			this.colAccessory.AppearanceHeader.Options.UseBackColor = true;
			this.colAccessory.AppearanceHeader.Options.UseBorderColor = true;
			this.colAccessory.AppearanceHeader.Options.UseFont = true;
			this.colAccessory.AppearanceHeader.Options.UseForeColor = true;
			this.colAccessory.AppearanceHeader.Options.UseTextOptions = true;
			this.colAccessory.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colAccessory.Caption = "Accessory";
			this.colAccessory.ColumnEdit = this.repositoryItemMemoEdit1;
			this.colAccessory.FieldName = "Accessory";
			this.colAccessory.Name = "colAccessory";
			this.colAccessory.OptionsColumn.AllowEdit = false;
			this.colAccessory.Visible = true;
			this.colAccessory.VisibleIndex = 10;
			this.colAccessory.Width = 213;
			// 
			// colDescription
			// 
			this.colDescription.AppearanceCell.BackColor = System.Drawing.Color.White;
			this.colDescription.AppearanceCell.BackColor2 = System.Drawing.Color.White;
			this.colDescription.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.colDescription.AppearanceCell.Options.UseBackColor = true;
			this.colDescription.AppearanceCell.Options.UseFont = true;
			this.colDescription.AppearanceCell.Options.UseTextOptions = true;
			this.colDescription.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colDescription.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colDescription.AppearanceHeader.BackColor = System.Drawing.Color.DodgerBlue;
			this.colDescription.AppearanceHeader.BackColor2 = System.Drawing.Color.DodgerBlue;
			this.colDescription.AppearanceHeader.BorderColor = System.Drawing.Color.DodgerBlue;
			this.colDescription.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colDescription.AppearanceHeader.ForeColor = System.Drawing.Color.White;
			this.colDescription.AppearanceHeader.Options.UseBackColor = true;
			this.colDescription.AppearanceHeader.Options.UseBorderColor = true;
			this.colDescription.AppearanceHeader.Options.UseFont = true;
			this.colDescription.AppearanceHeader.Options.UseForeColor = true;
			this.colDescription.AppearanceHeader.Options.UseTextOptions = true;
			this.colDescription.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colDescription.Caption = "Mô tả";
			this.colDescription.ColumnEdit = this.repositoryItemMemoEdit1;
			this.colDescription.FieldName = "Description";
			this.colDescription.Name = "colDescription";
			this.colDescription.OptionsColumn.AllowEdit = false;
			this.colDescription.Visible = true;
			this.colDescription.VisibleIndex = 1;
			this.colDescription.Width = 207;
			// 
			// colCreatDate
			// 
			this.colCreatDate.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.colCreatDate.AppearanceCell.Options.UseFont = true;
			this.colCreatDate.AppearanceCell.Options.UseTextOptions = true;
			this.colCreatDate.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colCreatDate.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colCreatDate.AppearanceHeader.BackColor = System.Drawing.Color.DodgerBlue;
			this.colCreatDate.AppearanceHeader.BackColor2 = System.Drawing.Color.DodgerBlue;
			this.colCreatDate.AppearanceHeader.BorderColor = System.Drawing.Color.DodgerBlue;
			this.colCreatDate.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colCreatDate.AppearanceHeader.ForeColor = System.Drawing.Color.White;
			this.colCreatDate.AppearanceHeader.Options.UseBackColor = true;
			this.colCreatDate.AppearanceHeader.Options.UseBorderColor = true;
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
			this.colCreatDate.VisibleIndex = 12;
			this.colCreatDate.Width = 117;
			// 
			// colQty
			// 
			this.colQty.AppearanceCell.BackColor = System.Drawing.Color.White;
			this.colQty.AppearanceCell.BackColor2 = System.Drawing.Color.White;
			this.colQty.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.colQty.AppearanceCell.Options.UseBackColor = true;
			this.colQty.AppearanceCell.Options.UseFont = true;
			this.colQty.AppearanceCell.Options.UseTextOptions = true;
			this.colQty.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colQty.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colQty.AppearanceHeader.BackColor = System.Drawing.Color.DodgerBlue;
			this.colQty.AppearanceHeader.BackColor2 = System.Drawing.Color.DodgerBlue;
			this.colQty.AppearanceHeader.BorderColor = System.Drawing.Color.DodgerBlue;
			this.colQty.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colQty.AppearanceHeader.ForeColor = System.Drawing.Color.White;
			this.colQty.AppearanceHeader.Options.UseBackColor = true;
			this.colQty.AppearanceHeader.Options.UseBorderColor = true;
			this.colQty.AppearanceHeader.Options.UseFont = true;
			this.colQty.AppearanceHeader.Options.UseForeColor = true;
			this.colQty.AppearanceHeader.Options.UseTextOptions = true;
			this.colQty.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colQty.Caption = "Q\'TY";
			this.colQty.FieldName = "Qty";
			this.colQty.Name = "colQty";
			this.colQty.OptionsColumn.AllowEdit = false;
			this.colQty.Visible = true;
			this.colQty.VisibleIndex = 3;
			this.colQty.Width = 103;
			// 
			// colRealQty
			// 
			this.colRealQty.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.colRealQty.AppearanceCell.Options.UseFont = true;
			this.colRealQty.AppearanceCell.Options.UseTextOptions = true;
			this.colRealQty.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colRealQty.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colRealQty.AppearanceHeader.BackColor = System.Drawing.Color.DodgerBlue;
			this.colRealQty.AppearanceHeader.BackColor2 = System.Drawing.Color.DodgerBlue;
			this.colRealQty.AppearanceHeader.BorderColor = System.Drawing.Color.DodgerBlue;
			this.colRealQty.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colRealQty.AppearanceHeader.ForeColor = System.Drawing.Color.White;
			this.colRealQty.AppearanceHeader.Options.UseBackColor = true;
			this.colRealQty.AppearanceHeader.Options.UseBorderColor = true;
			this.colRealQty.AppearanceHeader.Options.UseFont = true;
			this.colRealQty.AppearanceHeader.Options.UseForeColor = true;
			this.colRealQty.AppearanceHeader.Options.UseTextOptions = true;
			this.colRealQty.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colRealQty.Caption = "Real Q\'TY";
			this.colRealQty.FieldName = "RealQty";
			this.colRealQty.Name = "colRealQty";
			this.colRealQty.Visible = true;
			this.colRealQty.VisibleIndex = 11;
			this.colRealQty.Width = 98;
			// 
			// colStockCDID
			// 
			this.colStockCDID.Caption = "StockCDID";
			this.colStockCDID.FieldName = "StockCDID";
			this.colStockCDID.Name = "colStockCDID";
			// 
			// col1
			// 
			this.col1.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.col1.AppearanceCell.Options.UseFont = true;
			this.col1.AppearanceCell.Options.UseTextOptions = true;
			this.col1.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.col1.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.col1.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.col1.AppearanceHeader.BackColor = System.Drawing.Color.DodgerBlue;
			this.col1.AppearanceHeader.BackColor2 = System.Drawing.Color.DodgerBlue;
			this.col1.AppearanceHeader.BorderColor = System.Drawing.Color.DodgerBlue;
			this.col1.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.col1.AppearanceHeader.ForeColor = System.Drawing.Color.White;
			this.col1.AppearanceHeader.Options.UseBackColor = true;
			this.col1.AppearanceHeader.Options.UseBorderColor = true;
			this.col1.AppearanceHeader.Options.UseFont = true;
			this.col1.AppearanceHeader.Options.UseForeColor = true;
			this.col1.AppearanceHeader.Options.UseTextOptions = true;
			this.col1.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.col1.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.col1.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.col1.Caption = "1";
			this.col1.ColumnEdit = this.repositoryItemMemoEdit1;
			this.col1.FieldName = "Column1";
			this.col1.Name = "col1";
			this.col1.Visible = true;
			this.col1.VisibleIndex = 5;
			this.col1.Width = 138;
			// 
			// col2
			// 
			this.col2.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.col2.AppearanceCell.Options.UseFont = true;
			this.col2.AppearanceCell.Options.UseTextOptions = true;
			this.col2.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.col2.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.col2.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.col2.AppearanceHeader.BackColor = System.Drawing.Color.DodgerBlue;
			this.col2.AppearanceHeader.BackColor2 = System.Drawing.Color.DodgerBlue;
			this.col2.AppearanceHeader.BorderColor = System.Drawing.Color.DodgerBlue;
			this.col2.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.col2.AppearanceHeader.ForeColor = System.Drawing.Color.White;
			this.col2.AppearanceHeader.Options.UseBackColor = true;
			this.col2.AppearanceHeader.Options.UseBorderColor = true;
			this.col2.AppearanceHeader.Options.UseFont = true;
			this.col2.AppearanceHeader.Options.UseForeColor = true;
			this.col2.AppearanceHeader.Options.UseTextOptions = true;
			this.col2.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.col2.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.col2.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.col2.Caption = "2";
			this.col2.ColumnEdit = this.repositoryItemMemoEdit1;
			this.col2.FieldName = "Column2";
			this.col2.Name = "col2";
			this.col2.Visible = true;
			this.col2.VisibleIndex = 6;
			this.col2.Width = 138;
			// 
			// col3
			// 
			this.col3.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.col3.AppearanceCell.Options.UseFont = true;
			this.col3.AppearanceCell.Options.UseTextOptions = true;
			this.col3.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.col3.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.col3.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.col3.AppearanceHeader.BackColor = System.Drawing.Color.DodgerBlue;
			this.col3.AppearanceHeader.BackColor2 = System.Drawing.Color.DodgerBlue;
			this.col3.AppearanceHeader.BorderColor = System.Drawing.Color.DodgerBlue;
			this.col3.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.col3.AppearanceHeader.ForeColor = System.Drawing.Color.White;
			this.col3.AppearanceHeader.Options.UseBackColor = true;
			this.col3.AppearanceHeader.Options.UseBorderColor = true;
			this.col3.AppearanceHeader.Options.UseFont = true;
			this.col3.AppearanceHeader.Options.UseForeColor = true;
			this.col3.AppearanceHeader.Options.UseTextOptions = true;
			this.col3.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.col3.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.col3.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.col3.Caption = "3";
			this.col3.ColumnEdit = this.repositoryItemMemoEdit1;
			this.col3.FieldName = "Column3";
			this.col3.Name = "col3";
			this.col3.Visible = true;
			this.col3.VisibleIndex = 7;
			this.col3.Width = 138;
			// 
			// col4
			// 
			this.col4.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.col4.AppearanceCell.Options.UseFont = true;
			this.col4.AppearanceCell.Options.UseTextOptions = true;
			this.col4.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.col4.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.col4.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.col4.AppearanceHeader.BackColor = System.Drawing.Color.DodgerBlue;
			this.col4.AppearanceHeader.BackColor2 = System.Drawing.Color.DodgerBlue;
			this.col4.AppearanceHeader.BorderColor = System.Drawing.Color.DodgerBlue;
			this.col4.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.col4.AppearanceHeader.ForeColor = System.Drawing.Color.White;
			this.col4.AppearanceHeader.Options.UseBackColor = true;
			this.col4.AppearanceHeader.Options.UseBorderColor = true;
			this.col4.AppearanceHeader.Options.UseFont = true;
			this.col4.AppearanceHeader.Options.UseForeColor = true;
			this.col4.AppearanceHeader.Options.UseTextOptions = true;
			this.col4.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.col4.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.col4.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.col4.Caption = "4";
			this.col4.ColumnEdit = this.repositoryItemMemoEdit1;
			this.col4.FieldName = "Column4";
			this.col4.Name = "col4";
			this.col4.Visible = true;
			this.col4.VisibleIndex = 8;
			this.col4.Width = 138;
			// 
			// col5
			// 
			this.col5.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.col5.AppearanceCell.Options.UseFont = true;
			this.col5.AppearanceCell.Options.UseTextOptions = true;
			this.col5.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.col5.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.col5.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.col5.AppearanceHeader.BackColor = System.Drawing.Color.DodgerBlue;
			this.col5.AppearanceHeader.BackColor2 = System.Drawing.Color.DodgerBlue;
			this.col5.AppearanceHeader.BorderColor = System.Drawing.Color.DodgerBlue;
			this.col5.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.col5.AppearanceHeader.ForeColor = System.Drawing.Color.White;
			this.col5.AppearanceHeader.Options.UseBackColor = true;
			this.col5.AppearanceHeader.Options.UseBorderColor = true;
			this.col5.AppearanceHeader.Options.UseFont = true;
			this.col5.AppearanceHeader.Options.UseForeColor = true;
			this.col5.AppearanceHeader.Options.UseTextOptions = true;
			this.col5.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.col5.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.col5.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.col5.Caption = "5";
			this.col5.ColumnEdit = this.repositoryItemMemoEdit1;
			this.col5.FieldName = "Column5";
			this.col5.Name = "col5";
			this.col5.Visible = true;
			this.col5.VisibleIndex = 9;
			this.col5.Width = 138;
			// 
			// frmHistoryPartOutDetail
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(1024, 536);
			this.Controls.Add(this.grdData);
			this.Name = "frmHistoryPartOutDetail";
			this.Text = "CHI TIẾT KIỂM TRA";
			this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
			this.Load += new System.EventHandler(this.frmHistoryCheckDetail_Load);
			((System.ComponentModel.ISupportInitialize)(this.grdData)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grvData)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemMemoEdit1)).EndInit();
			this.ResumeLayout(false);

        }

        #endregion
        private DevExpress.XtraGrid.Columns.GridColumn gridColumn1;
        private DevExpress.XtraGrid.Columns.GridColumn colValueTypeText;
        private DevExpress.XtraGrid.Columns.GridColumn colRealValue1;
		private DevExpress.XtraGrid.GridControl grdData;
		private DevExpress.XtraGrid.Views.Grid.GridView grvData;
		private DevExpress.XtraGrid.Columns.GridColumn colID;
		private DevExpress.XtraGrid.Columns.GridColumn colStockID;
		private DevExpress.XtraGrid.Columns.GridColumn colShelf;
		private DevExpress.XtraGrid.Columns.GridColumn colLo;
		private DevExpress.XtraGrid.Columns.GridColumn colArticleID;
		private DevExpress.XtraGrid.Columns.GridColumn colAccessory;
		private DevExpress.XtraGrid.Columns.GridColumn colDescription;
		private DevExpress.XtraGrid.Columns.GridColumn colCreatDate;
		private DevExpress.XtraGrid.Columns.GridColumn colQty;
		private DevExpress.XtraGrid.Columns.GridColumn colRealQty;
		private DevExpress.XtraGrid.Columns.GridColumn colStockCDID;
		private DevExpress.XtraGrid.Columns.GridColumn col1;
		private DevExpress.XtraGrid.Columns.GridColumn col2;
		private DevExpress.XtraGrid.Columns.GridColumn col3;
		private DevExpress.XtraGrid.Columns.GridColumn col4;
		private DevExpress.XtraGrid.Columns.GridColumn col5;
		private DevExpress.XtraEditors.Repository.RepositoryItemMemoEdit repositoryItemMemoEdit1;
	}
}