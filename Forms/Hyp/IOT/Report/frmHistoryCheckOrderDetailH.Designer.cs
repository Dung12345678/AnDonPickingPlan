
namespace BMS
{
	partial class frmHistoryCheckOrderDetailH
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
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmHistoryCheckOrderDetailH));
			this.btnExportExecl = new DevExpress.XtraEditors.SimpleButton();
			this.grdData = new DevExpress.XtraGrid.GridControl();
			this.grvData = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.colPID = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colSTT = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colCreatedDate = new DevExpress.XtraGrid.Columns.GridColumn();
			this.repositoryItemMemoEdit1 = new DevExpress.XtraEditors.Repository.RepositoryItemMemoEdit();
			this.colGroup = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colID = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colDes = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colPath = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colNote = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colS = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colInFilm = new DevExpress.XtraGrid.Columns.GridColumn();
			this.repositoryItemCheckEdit1 = new DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit();
			this.colUpdatedDate = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colPrice = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colOrderTime = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colTem = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colTK = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colRealValue1 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colRealValue2 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colRealValue3 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colRealValue4 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colRealValue5 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.repositoryItemCheckEdit2 = new DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit();
			((System.ComponentModel.ISupportInitialize)(this.grdData)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grvData)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemMemoEdit1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemCheckEdit1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemCheckEdit2)).BeginInit();
			this.SuspendLayout();
			// 
			// btnExportExecl
			// 
			this.btnExportExecl.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExportExecl.Appearance.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.btnExportExecl.Appearance.Options.UseFont = true;
			this.btnExportExecl.Image = ((System.Drawing.Image)(resources.GetObject("btnExportExecl.Image")));
			this.btnExportExecl.Location = new System.Drawing.Point(1034, 12);
			this.btnExportExecl.Name = "btnExportExecl";
			this.btnExportExecl.Size = new System.Drawing.Size(89, 32);
			this.btnExportExecl.TabIndex = 29;
			this.btnExportExecl.Text = "Xuất excel";
			this.btnExportExecl.Click += new System.EventHandler(this.btnExportExecl_Click);
			// 
			// grdData
			// 
			this.grdData.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.grdData.Font = new System.Drawing.Font("Times New Roman", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.grdData.Location = new System.Drawing.Point(2, 3);
			this.grdData.MainView = this.grvData;
			this.grdData.Name = "grdData";
			this.grdData.RepositoryItems.AddRange(new DevExpress.XtraEditors.Repository.RepositoryItem[] {
            this.repositoryItemCheckEdit1,
            this.repositoryItemMemoEdit1,
            this.repositoryItemCheckEdit2});
			this.grdData.Size = new System.Drawing.Size(1141, 531);
			this.grdData.TabIndex = 30;
			this.grdData.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.grvData});
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
            this.colPID,
            this.colSTT,
            this.colCreatedDate,
            this.colGroup,
            this.colID,
            this.colDes,
            this.colPath,
            this.colNote,
            this.colS,
            this.colInFilm,
            this.colUpdatedDate,
            this.colPrice,
            this.colOrderTime,
            this.colTem,
            this.colTK,
            this.colRealValue1,
            this.colRealValue2,
            this.colRealValue3,
            this.colRealValue4,
            this.colRealValue5});
			this.grvData.GridControl = this.grdData;
			this.grvData.GroupCount = 1;
			this.grvData.HorzScrollStep = 5;
			this.grvData.Name = "grvData";
			this.grvData.OptionsBehavior.AutoExpandAllGroups = true;
			this.grvData.OptionsCustomization.AllowRowSizing = true;
			this.grvData.OptionsFind.AlwaysVisible = true;
			this.grvData.OptionsFind.ShowCloseButton = false;
			this.grvData.OptionsSelection.MultiSelect = true;
			this.grvData.OptionsView.RowAutoHeight = true;
			this.grvData.OptionsView.ShowAutoFilterRow = true;
			this.grvData.OptionsView.ShowGroupPanel = false;
			this.grvData.SortInfo.AddRange(new DevExpress.XtraGrid.Columns.GridColumnSortInfo[] {
            new DevExpress.XtraGrid.Columns.GridColumnSortInfo(this.colGroup, DevExpress.Data.ColumnSortOrder.Ascending)});
			// 
			// colPID
			// 
			this.colPID.AppearanceCell.Options.UseTextOptions = true;
			this.colPID.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colPID.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colPID.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colPID.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colPID.AppearanceHeader.Options.UseFont = true;
			this.colPID.AppearanceHeader.Options.UseTextOptions = true;
			this.colPID.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colPID.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colPID.Caption = "Giá trị tiêu chuẩn";
			this.colPID.FieldName = "StandardValue";
			this.colPID.Name = "colPID";
			this.colPID.OptionsColumn.AllowEdit = false;
			this.colPID.OptionsColumn.AllowSize = false;
			this.colPID.Visible = true;
			this.colPID.VisibleIndex = 3;
			this.colPID.Width = 182;
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
			this.colSTT.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colSTT.Caption = "STT";
			this.colSTT.FieldName = "RowNum";
			this.colSTT.Name = "colSTT";
			this.colSTT.OptionsColumn.AllowEdit = false;
			this.colSTT.OptionsColumn.AllowSize = false;
			this.colSTT.Summary.AddRange(new DevExpress.XtraGrid.GridSummaryItem[] {
            new DevExpress.XtraGrid.GridColumnSummaryItem(DevExpress.Data.SummaryItemType.Count)});
			this.colSTT.Visible = true;
			this.colSTT.VisibleIndex = 0;
			this.colSTT.Width = 53;
			// 
			// colCreatedDate
			// 
			this.colCreatedDate.AppearanceCell.Options.UseTextOptions = true;
			this.colCreatedDate.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colCreatedDate.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colCreatedDate.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colCreatedDate.AppearanceHeader.Options.UseFont = true;
			this.colCreatedDate.AppearanceHeader.Options.UseTextOptions = true;
			this.colCreatedDate.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colCreatedDate.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colCreatedDate.Caption = "Ghi chép";
			this.colCreatedDate.ColumnEdit = this.repositoryItemMemoEdit1;
			this.colCreatedDate.FieldName = "ValueTypeText";
			this.colCreatedDate.Name = "colCreatedDate";
			this.colCreatedDate.OptionsColumn.AllowEdit = false;
			this.colCreatedDate.OptionsColumn.AllowSize = false;
			this.colCreatedDate.Visible = true;
			this.colCreatedDate.VisibleIndex = 2;
			this.colCreatedDate.Width = 93;
			// 
			// repositoryItemMemoEdit1
			// 
			this.repositoryItemMemoEdit1.Appearance.Options.UseTextOptions = true;
			this.repositoryItemMemoEdit1.Appearance.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.repositoryItemMemoEdit1.Name = "repositoryItemMemoEdit1";
			// 
			// colGroup
			// 
			this.colGroup.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colGroup.AppearanceCell.Options.UseFont = true;
			this.colGroup.AppearanceCell.Options.UseTextOptions = true;
			this.colGroup.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colGroup.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colGroup.AppearanceHeader.Options.UseFont = true;
			this.colGroup.AppearanceHeader.Options.UseTextOptions = true;
			this.colGroup.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colGroup.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colGroup.Caption = "Công đoạn";
			this.colGroup.FieldName = "ProductStepInfo";
			this.colGroup.Name = "colGroup";
			this.colGroup.OptionsColumn.AllowEdit = false;
			// 
			// colID
			// 
			this.colID.AppearanceCell.Options.UseTextOptions = true;
			this.colID.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colID.AppearanceHeader.Options.UseTextOptions = true;
			this.colID.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colID.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colID.Caption = "ID";
			this.colID.FieldName = "ID";
			this.colID.Name = "colID";
			this.colID.OptionsColumn.AllowEdit = false;
			// 
			// colDes
			// 
			this.colDes.AppearanceCell.Options.UseTextOptions = true;
			this.colDes.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colDes.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colDes.AppearanceHeader.Options.UseFont = true;
			this.colDes.AppearanceHeader.Options.UseTextOptions = true;
			this.colDes.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colDes.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colDes.Caption = "Mô tả";
			this.colDes.FieldName = "Description";
			this.colDes.Name = "colDes";
			this.colDes.OptionsColumn.AllowEdit = false;
			this.colDes.Width = 141;
			// 
			// colPath
			// 
			this.colPath.AppearanceCell.Options.UseTextOptions = true;
			this.colPath.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colPath.AppearanceHeader.Options.UseTextOptions = true;
			this.colPath.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colPath.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colPath.Caption = "Đường dẫn";
			this.colPath.FieldName = "Path";
			this.colPath.Name = "colPath";
			this.colPath.OptionsColumn.AllowEdit = false;
			this.colPath.Width = 104;
			// 
			// colNote
			// 
			this.colNote.AppearanceCell.Options.UseTextOptions = true;
			this.colNote.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colNote.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colNote.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colNote.AppearanceHeader.Options.UseFont = true;
			this.colNote.AppearanceHeader.Options.UseTextOptions = true;
			this.colNote.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colNote.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colNote.Caption = "Mục kiểm tra";
			this.colNote.ColumnEdit = this.repositoryItemMemoEdit1;
			this.colNote.FieldName = "ProductWorkingName";
			this.colNote.Name = "colNote";
			this.colNote.OptionsColumn.AllowEdit = false;
			this.colNote.OptionsColumn.AllowSize = false;
			this.colNote.Visible = true;
			this.colNote.VisibleIndex = 1;
			this.colNote.Width = 271;
			// 
			// colS
			// 
			this.colS.AppearanceCell.Options.UseTextOptions = true;
			this.colS.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colS.AppearanceHeader.Options.UseTextOptions = true;
			this.colS.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colS.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colS.Caption = "Trạng thái";
			this.colS.FieldName = "Status";
			this.colS.Name = "colS";
			this.colS.OptionsColumn.AllowEdit = false;
			// 
			// colInFilm
			// 
			this.colInFilm.AppearanceCell.Options.UseTextOptions = true;
			this.colInFilm.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colInFilm.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colInFilm.AppearanceHeader.Options.UseFont = true;
			this.colInFilm.AppearanceHeader.Options.UseTextOptions = true;
			this.colInFilm.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colInFilm.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colInFilm.Caption = "Film in";
			this.colInFilm.ColumnEdit = this.repositoryItemCheckEdit1;
			this.colInFilm.FieldName = "HasFilm";
			this.colInFilm.Name = "colInFilm";
			this.colInFilm.OptionsColumn.AllowEdit = false;
			this.colInFilm.OptionsColumn.AllowSize = false;
			this.colInFilm.Width = 31;
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
			// colUpdatedDate
			// 
			this.colUpdatedDate.AppearanceCell.Options.UseTextOptions = true;
			this.colUpdatedDate.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colUpdatedDate.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colUpdatedDate.AppearanceHeader.Options.UseFont = true;
			this.colUpdatedDate.AppearanceHeader.Options.UseTextOptions = true;
			this.colUpdatedDate.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colUpdatedDate.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colUpdatedDate.Caption = "Ngày sửa";
			this.colUpdatedDate.FieldName = "UpdatedDate";
			this.colUpdatedDate.Name = "colUpdatedDate";
			this.colUpdatedDate.OptionsColumn.AllowEdit = false;
			// 
			// colPrice
			// 
			this.colPrice.AppearanceCell.Options.UseTextOptions = true;
			this.colPrice.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colPrice.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colPrice.AppearanceHeader.Options.UseFont = true;
			this.colPrice.AppearanceHeader.Options.UseTextOptions = true;
			this.colPrice.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colPrice.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colPrice.Caption = "Giá(VND)";
			this.colPrice.DisplayFormat.FormatString = "n2";
			this.colPrice.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.colPrice.FieldName = "Price";
			this.colPrice.Name = "colPrice";
			this.colPrice.OptionsColumn.AllowEdit = false;
			// 
			// colOrderTime
			// 
			this.colOrderTime.AppearanceCell.Options.UseTextOptions = true;
			this.colOrderTime.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colOrderTime.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colOrderTime.AppearanceHeader.Options.UseFont = true;
			this.colOrderTime.AppearanceHeader.Options.UseTextOptions = true;
			this.colOrderTime.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colOrderTime.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colOrderTime.Caption = "Thời gian đặt hàng(ngày)";
			this.colOrderTime.DisplayFormat.FormatString = "n2";
			this.colOrderTime.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.colOrderTime.FieldName = "OrderTime";
			this.colOrderTime.Name = "colOrderTime";
			this.colOrderTime.OptionsColumn.AllowEdit = false;
			// 
			// colTem
			// 
			this.colTem.AppearanceCell.Options.UseTextOptions = true;
			this.colTem.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colTem.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colTem.AppearanceHeader.Options.UseFont = true;
			this.colTem.AppearanceHeader.Options.UseTextOptions = true;
			this.colTem.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colTem.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colTem.Caption = "Tem";
			this.colTem.ColumnEdit = this.repositoryItemCheckEdit1;
			this.colTem.FieldName = "HasTem";
			this.colTem.Name = "colTem";
			this.colTem.OptionsColumn.AllowEdit = false;
			this.colTem.OptionsColumn.AllowSize = false;
			this.colTem.Width = 38;
			// 
			// colTK
			// 
			this.colTK.AppearanceCell.Options.UseTextOptions = true;
			this.colTK.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colTK.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colTK.AppearanceHeader.Options.UseFont = true;
			this.colTK.AppearanceHeader.Options.UseTextOptions = true;
			this.colTK.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colTK.Caption = "TK";
			this.colTK.ColumnEdit = this.repositoryItemCheckEdit1;
			this.colTK.FieldName = "HasTK";
			this.colTK.Name = "colTK";
			this.colTK.OptionsColumn.AllowEdit = false;
			this.colTK.OptionsColumn.AllowSize = false;
			this.colTK.UnboundType = DevExpress.Data.UnboundColumnType.Boolean;
			this.colTK.Width = 35;
			// 
			// colRealValue1
			// 
			this.colRealValue1.AppearanceCell.Options.UseTextOptions = true;
			this.colRealValue1.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colRealValue1.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colRealValue1.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colRealValue1.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.colRealValue1.AppearanceHeader.Options.UseFont = true;
			this.colRealValue1.AppearanceHeader.Options.UseTextOptions = true;
			this.colRealValue1.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colRealValue1.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colRealValue1.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colRealValue1.Caption = "#";
			this.colRealValue1.Name = "colRealValue1";
			this.colRealValue1.OptionsColumn.AllowEdit = false;
			this.colRealValue1.Visible = true;
			this.colRealValue1.VisibleIndex = 4;
			this.colRealValue1.Width = 120;
			// 
			// colRealValue2
			// 
			this.colRealValue2.AppearanceCell.Options.UseTextOptions = true;
			this.colRealValue2.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colRealValue2.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colRealValue2.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colRealValue2.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.colRealValue2.AppearanceHeader.Options.UseFont = true;
			this.colRealValue2.AppearanceHeader.Options.UseTextOptions = true;
			this.colRealValue2.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colRealValue2.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colRealValue2.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colRealValue2.Caption = "#";
			this.colRealValue2.Name = "colRealValue2";
			this.colRealValue2.OptionsColumn.AllowEdit = false;
			this.colRealValue2.Visible = true;
			this.colRealValue2.VisibleIndex = 5;
			this.colRealValue2.Width = 114;
			// 
			// colRealValue3
			// 
			this.colRealValue3.AppearanceCell.Options.UseTextOptions = true;
			this.colRealValue3.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colRealValue3.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colRealValue3.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colRealValue3.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.colRealValue3.AppearanceHeader.Options.UseFont = true;
			this.colRealValue3.AppearanceHeader.Options.UseTextOptions = true;
			this.colRealValue3.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colRealValue3.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colRealValue3.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colRealValue3.Caption = "#";
			this.colRealValue3.Name = "colRealValue3";
			this.colRealValue3.OptionsColumn.AllowEdit = false;
			this.colRealValue3.Visible = true;
			this.colRealValue3.VisibleIndex = 6;
			this.colRealValue3.Width = 106;
			// 
			// colRealValue4
			// 
			this.colRealValue4.AppearanceCell.Options.UseTextOptions = true;
			this.colRealValue4.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colRealValue4.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colRealValue4.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colRealValue4.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.colRealValue4.AppearanceHeader.Options.UseFont = true;
			this.colRealValue4.AppearanceHeader.Options.UseTextOptions = true;
			this.colRealValue4.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colRealValue4.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colRealValue4.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colRealValue4.Caption = "#";
			this.colRealValue4.Name = "colRealValue4";
			this.colRealValue4.OptionsColumn.AllowEdit = false;
			this.colRealValue4.Visible = true;
			this.colRealValue4.VisibleIndex = 7;
			this.colRealValue4.Width = 97;
			// 
			// colRealValue5
			// 
			this.colRealValue5.AppearanceCell.Options.UseTextOptions = true;
			this.colRealValue5.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colRealValue5.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colRealValue5.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colRealValue5.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.colRealValue5.AppearanceHeader.Options.UseFont = true;
			this.colRealValue5.AppearanceHeader.Options.UseTextOptions = true;
			this.colRealValue5.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colRealValue5.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colRealValue5.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colRealValue5.Caption = "#";
			this.colRealValue5.Name = "colRealValue5";
			this.colRealValue5.OptionsColumn.AllowEdit = false;
			this.colRealValue5.Visible = true;
			this.colRealValue5.VisibleIndex = 8;
			this.colRealValue5.Width = 87;
			// 
			// repositoryItemCheckEdit2
			// 
			this.repositoryItemCheckEdit2.AutoHeight = false;
			this.repositoryItemCheckEdit2.Caption = "Check";
			this.repositoryItemCheckEdit2.Name = "repositoryItemCheckEdit2";
			this.repositoryItemCheckEdit2.NullStyle = DevExpress.XtraEditors.Controls.StyleIndeterminate.Unchecked;
			// 
			// frmHistoryCheckOrderDetailH
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(1145, 536);
			this.Controls.Add(this.btnExportExecl);
			this.Controls.Add(this.grdData);
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.Name = "frmHistoryCheckOrderDetailH";
			this.Text = "CHI TIẾT KIỂM TRA THEO ORDER HYP";
			this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
			this.Load += new System.EventHandler(this.frmHistoryCheckOrderDetailH_Load);
			((System.ComponentModel.ISupportInitialize)(this.grdData)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grvData)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemMemoEdit1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemCheckEdit1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemCheckEdit2)).EndInit();
			this.ResumeLayout(false);

		}

		#endregion
		private DevExpress.XtraEditors.SimpleButton btnExportExecl;
		private DevExpress.XtraGrid.GridControl grdData;
		private DevExpress.XtraGrid.Views.Grid.GridView grvData;
		private DevExpress.XtraGrid.Columns.GridColumn colPID;
		private DevExpress.XtraGrid.Columns.GridColumn colSTT;
		private DevExpress.XtraGrid.Columns.GridColumn colCreatedDate;
		private DevExpress.XtraEditors.Repository.RepositoryItemMemoEdit repositoryItemMemoEdit1;
		private DevExpress.XtraGrid.Columns.GridColumn colGroup;
		private DevExpress.XtraGrid.Columns.GridColumn colID;
		private DevExpress.XtraGrid.Columns.GridColumn colDes;
		private DevExpress.XtraGrid.Columns.GridColumn colPath;
		private DevExpress.XtraGrid.Columns.GridColumn colNote;
		private DevExpress.XtraGrid.Columns.GridColumn colS;
		private DevExpress.XtraGrid.Columns.GridColumn colInFilm;
		private DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit repositoryItemCheckEdit1;
		private DevExpress.XtraGrid.Columns.GridColumn colUpdatedDate;
		private DevExpress.XtraGrid.Columns.GridColumn colPrice;
		private DevExpress.XtraGrid.Columns.GridColumn colOrderTime;
		private DevExpress.XtraGrid.Columns.GridColumn colTem;
		private DevExpress.XtraGrid.Columns.GridColumn colTK;
		private DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit repositoryItemCheckEdit2;
		private DevExpress.XtraGrid.Columns.GridColumn colRealValue1;
		private DevExpress.XtraGrid.Columns.GridColumn colRealValue2;
		private DevExpress.XtraGrid.Columns.GridColumn colRealValue3;
		private DevExpress.XtraGrid.Columns.GridColumn colRealValue4;
		private DevExpress.XtraGrid.Columns.GridColumn colRealValue5;
	}
}