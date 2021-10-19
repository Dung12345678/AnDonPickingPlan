
namespace BMS
{
    partial class frmAssemblyOrder
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
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmAssemblyOrder));
			this.toolStrip1 = new System.Windows.Forms.ToolStrip();
			this.btnCreatePlan = new System.Windows.Forms.ToolStripButton();
			this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
			this.btnEditPlan = new System.Windows.Forms.ToolStripButton();
			this.toolStripSeparator7 = new System.Windows.Forms.ToolStripSeparator();
			this.btnDelPlan = new System.Windows.Forms.ToolStripButton();
			this.toolStripSeparator8 = new System.Windows.Forms.ToolStripSeparator();
			this.btnExportExcel = new System.Windows.Forms.ToolStripButton();
			this.toolStripSeparator11 = new System.Windows.Forms.ToolStripSeparator();
			this.btnImportExcel = new System.Windows.Forms.ToolStripButton();
			this.toolStripSeparator12 = new System.Windows.Forms.ToolStripSeparator();
			this.grdData = new DevExpress.XtraGrid.GridControl();
			this.grvData = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.colID = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colAssemblyOrder = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colArticleID = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colCnt = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colDescription = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colQty = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colShelf = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colLocation = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colOrderCode = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colLot = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colCreateAt = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colUserr = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colPID = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colOrderQty = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colProductGroupCode = new DevExpress.XtraGrid.Columns.GridColumn();
			this.timer1 = new System.Windows.Forms.Timer(this.components);
			this.panel2 = new System.Windows.Forms.Panel();
			this.btnSearch = new DevExpress.XtraEditors.SimpleButton();
			this.txbSearch = new System.Windows.Forms.TextBox();
			this.dtpFrom = new System.Windows.Forms.DateTimePicker();
			this.dtpTo = new System.Windows.Forms.DateTimePicker();
			this.label1 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.chkExportExcelAltax = new System.Windows.Forms.CheckBox();
			this.chkExportExcelHyp = new System.Windows.Forms.CheckBox();
			this.btnDeleteDuplicate = new DevExpress.XtraEditors.SimpleButton();
			this.toolStrip1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.grdData)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grvData)).BeginInit();
			this.panel2.SuspendLayout();
			this.SuspendLayout();
			// 
			// toolStrip1
			// 
			this.toolStrip1.AutoSize = false;
			this.toolStrip1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(234)))), ((int)(((byte)(234)))), ((int)(((byte)(234)))));
			this.toolStrip1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.toolStrip1.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden;
			this.toolStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
			this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.btnCreatePlan,
            this.toolStripSeparator1,
            this.btnEditPlan,
            this.toolStripSeparator7,
            this.btnDelPlan,
            this.toolStripSeparator8,
            this.btnExportExcel,
            this.toolStripSeparator11,
            this.btnImportExcel,
            this.toolStripSeparator12});
			this.toolStrip1.LayoutStyle = System.Windows.Forms.ToolStripLayoutStyle.Flow;
			this.toolStrip1.Location = new System.Drawing.Point(0, 0);
			this.toolStrip1.Name = "toolStrip1";
			this.toolStrip1.Padding = new System.Windows.Forms.Padding(0, 0, 2, 0);
			this.toolStrip1.RenderMode = System.Windows.Forms.ToolStripRenderMode.Professional;
			this.toolStrip1.Size = new System.Drawing.Size(954, 48);
			this.toolStrip1.Stretch = true;
			this.toolStrip1.TabIndex = 27;
			this.toolStrip1.Text = "toolStrip2";
			// 
			// btnCreatePlan
			// 
			this.btnCreatePlan.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnCreatePlan.Image = ((System.Drawing.Image)(resources.GetObject("btnCreatePlan.Image")));
			this.btnCreatePlan.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnCreatePlan.Name = "btnCreatePlan";
			this.btnCreatePlan.Size = new System.Drawing.Size(102, 42);
			this.btnCreatePlan.Tag = "frmProduct_AddProductH";
			this.btnCreatePlan.Text = "Thêm Order";
			this.btnCreatePlan.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnCreatePlan.Click += new System.EventHandler(this.btnCreatePlan_Click);
			// 
			// toolStripSeparator1
			// 
			this.toolStripSeparator1.Name = "toolStripSeparator1";
			this.toolStripSeparator1.Size = new System.Drawing.Size(6, 23);
			// 
			// btnEditPlan
			// 
			this.btnEditPlan.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnEditPlan.Image = ((System.Drawing.Image)(resources.GetObject("btnEditPlan.Image")));
			this.btnEditPlan.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnEditPlan.Name = "btnEditPlan";
			this.btnEditPlan.Size = new System.Drawing.Size(89, 42);
			this.btnEditPlan.Tag = "frmProduct_EditProductH";
			this.btnEditPlan.Text = "Sửa Order";
			this.btnEditPlan.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnEditPlan.Click += new System.EventHandler(this.btnEditPlan_Click);
			// 
			// toolStripSeparator7
			// 
			this.toolStripSeparator7.Name = "toolStripSeparator7";
			this.toolStripSeparator7.Size = new System.Drawing.Size(6, 23);
			// 
			// btnDelPlan
			// 
			this.btnDelPlan.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnDelPlan.Image = ((System.Drawing.Image)(resources.GetObject("btnDelPlan.Image")));
			this.btnDelPlan.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnDelPlan.Name = "btnDelPlan";
			this.btnDelPlan.Size = new System.Drawing.Size(90, 42);
			this.btnDelPlan.Tag = "frmProduct_DeleteProductH";
			this.btnDelPlan.Text = "Xóa Order";
			this.btnDelPlan.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnDelPlan.Click += new System.EventHandler(this.btnDelPlan_Click);
			// 
			// toolStripSeparator8
			// 
			this.toolStripSeparator8.AutoSize = false;
			this.toolStripSeparator8.Name = "toolStripSeparator8";
			this.toolStripSeparator8.Size = new System.Drawing.Size(6, 40);
			// 
			// btnExportExcel
			// 
			this.btnExportExcel.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnExportExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExportExcel.Image")));
			this.btnExportExcel.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnExportExcel.Name = "btnExportExcel";
			this.btnExportExcel.Size = new System.Drawing.Size(92, 42);
			this.btnExportExcel.Tag = "";
			this.btnExportExcel.Text = "Xuất Excel";
			this.btnExportExcel.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnExportExcel.Click += new System.EventHandler(this.btnExportExcel_Click);
			// 
			// toolStripSeparator11
			// 
			this.toolStripSeparator11.Name = "toolStripSeparator11";
			this.toolStripSeparator11.Size = new System.Drawing.Size(6, 23);
			// 
			// btnImportExcel
			// 
			this.btnImportExcel.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnImportExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnImportExcel.Image")));
			this.btnImportExcel.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnImportExcel.Name = "btnImportExcel";
			this.btnImportExcel.Size = new System.Drawing.Size(97, 42);
			this.btnImportExcel.Tag = "";
			this.btnImportExcel.Text = "Nhập Excel";
			this.btnImportExcel.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnImportExcel.Click += new System.EventHandler(this.btnImportExcel_Click);
			// 
			// toolStripSeparator12
			// 
			this.toolStripSeparator12.AutoSize = false;
			this.toolStripSeparator12.Name = "toolStripSeparator12";
			this.toolStripSeparator12.Size = new System.Drawing.Size(6, 40);
			// 
			// grdData
			// 
			this.grdData.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.grdData.EmbeddedNavigator.Margin = new System.Windows.Forms.Padding(4);
			this.grdData.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.grdData.Location = new System.Drawing.Point(0, 80);
			this.grdData.LookAndFeel.UseDefaultLookAndFeel = false;
			this.grdData.MainView = this.grvData;
			this.grdData.Margin = new System.Windows.Forms.Padding(4);
			this.grdData.Name = "grdData";
			this.grdData.Size = new System.Drawing.Size(954, 370);
			this.grdData.TabIndex = 29;
			this.grdData.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.grvData});
			// 
			// grvData
			// 
			this.grvData.ColumnPanelRowHeight = 60;
			this.grvData.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
            this.colID,
            this.colAssemblyOrder,
            this.colArticleID,
            this.colCnt,
            this.colDescription,
            this.colQty,
            this.colShelf,
            this.colLocation,
            this.colOrderCode,
            this.colLot,
            this.colCreateAt,
            this.colUserr,
            this.colPID,
            this.colOrderQty,
            this.colProductGroupCode});
			this.grvData.GridControl = this.grdData;
			this.grvData.HorzScrollVisibility = DevExpress.XtraGrid.Views.Base.ScrollVisibility.Always;
			this.grvData.Name = "grvData";
			this.grvData.OptionsBehavior.AllowAddRows = DevExpress.Utils.DefaultBoolean.False;
			this.grvData.OptionsBehavior.AllowDeleteRows = DevExpress.Utils.DefaultBoolean.False;
			this.grvData.OptionsBehavior.AllowFixedGroups = DevExpress.Utils.DefaultBoolean.True;
			this.grvData.OptionsBehavior.AllowIncrementalSearch = true;
			this.grvData.OptionsBehavior.AllowPartialRedrawOnScrolling = false;
			this.grvData.OptionsView.ColumnAutoWidth = false;
			this.grvData.OptionsView.ShowGroupPanel = false;
			this.grvData.RowHeight = 25;
			this.grvData.DoubleClick += new System.EventHandler(this.gvSonPlan_DoubleClick);
			// 
			// colID
			// 
			this.colID.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.colID.AppearanceCell.Options.UseFont = true;
			this.colID.AppearanceCell.Options.UseTextOptions = true;
			this.colID.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colID.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colID.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colID.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colID.AppearanceHeader.ForeColor = System.Drawing.Color.Black;
			this.colID.AppearanceHeader.Options.UseBackColor = true;
			this.colID.AppearanceHeader.Options.UseFont = true;
			this.colID.AppearanceHeader.Options.UseForeColor = true;
			this.colID.AppearanceHeader.Options.UseTextOptions = true;
			this.colID.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colID.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colID.Caption = "ID";
			this.colID.FieldName = "ID";
			this.colID.MinWidth = 25;
			this.colID.Name = "colID";
			this.colID.OptionsColumn.AllowEdit = false;
			// 
			// colAssemblyOrder
			// 
			this.colAssemblyOrder.AppearanceCell.BackColor = System.Drawing.Color.White;
			this.colAssemblyOrder.AppearanceCell.BackColor2 = System.Drawing.Color.White;
			this.colAssemblyOrder.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.colAssemblyOrder.AppearanceCell.Options.UseBackColor = true;
			this.colAssemblyOrder.AppearanceCell.Options.UseFont = true;
			this.colAssemblyOrder.AppearanceCell.Options.UseTextOptions = true;
			this.colAssemblyOrder.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colAssemblyOrder.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colAssemblyOrder.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colAssemblyOrder.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colAssemblyOrder.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colAssemblyOrder.AppearanceHeader.ForeColor = System.Drawing.Color.Black;
			this.colAssemblyOrder.AppearanceHeader.Options.UseBackColor = true;
			this.colAssemblyOrder.AppearanceHeader.Options.UseFont = true;
			this.colAssemblyOrder.AppearanceHeader.Options.UseForeColor = true;
			this.colAssemblyOrder.AppearanceHeader.Options.UseTextOptions = true;
			this.colAssemblyOrder.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colAssemblyOrder.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colAssemblyOrder.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colAssemblyOrder.Caption = "ASSEMBLY ORDER";
			this.colAssemblyOrder.FieldName = "OrderCode";
			this.colAssemblyOrder.MinWidth = 25;
			this.colAssemblyOrder.Name = "colAssemblyOrder";
			this.colAssemblyOrder.OptionsColumn.AllowEdit = false;
			this.colAssemblyOrder.Visible = true;
			this.colAssemblyOrder.VisibleIndex = 0;
			this.colAssemblyOrder.Width = 200;
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
			this.colArticleID.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colArticleID.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colArticleID.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colArticleID.AppearanceHeader.ForeColor = System.Drawing.Color.Black;
			this.colArticleID.AppearanceHeader.Options.UseBackColor = true;
			this.colArticleID.AppearanceHeader.Options.UseFont = true;
			this.colArticleID.AppearanceHeader.Options.UseForeColor = true;
			this.colArticleID.AppearanceHeader.Options.UseTextOptions = true;
			this.colArticleID.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colArticleID.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colArticleID.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colArticleID.Caption = "Article ID";
			this.colArticleID.FieldName = "ArticleID";
			this.colArticleID.MinWidth = 25;
			this.colArticleID.Name = "colArticleID";
			this.colArticleID.OptionsColumn.AllowEdit = false;
			this.colArticleID.Visible = true;
			this.colArticleID.VisibleIndex = 3;
			this.colArticleID.Width = 236;
			// 
			// colCnt
			// 
			this.colCnt.AppearanceCell.BackColor = System.Drawing.Color.White;
			this.colCnt.AppearanceCell.BackColor2 = System.Drawing.Color.White;
			this.colCnt.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.colCnt.AppearanceCell.Options.UseBackColor = true;
			this.colCnt.AppearanceCell.Options.UseFont = true;
			this.colCnt.AppearanceCell.Options.UseTextOptions = true;
			this.colCnt.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colCnt.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colCnt.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colCnt.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colCnt.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colCnt.AppearanceHeader.ForeColor = System.Drawing.Color.Black;
			this.colCnt.AppearanceHeader.Options.UseBackColor = true;
			this.colCnt.AppearanceHeader.Options.UseFont = true;
			this.colCnt.AppearanceHeader.Options.UseForeColor = true;
			this.colCnt.AppearanceHeader.Options.UseTextOptions = true;
			this.colCnt.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colCnt.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colCnt.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colCnt.Caption = "CNT";
			this.colCnt.FieldName = "Cnt";
			this.colCnt.MinWidth = 25;
			this.colCnt.Name = "colCnt";
			this.colCnt.OptionsColumn.AllowEdit = false;
			this.colCnt.Visible = true;
			this.colCnt.VisibleIndex = 4;
			this.colCnt.Width = 109;
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
			this.colDescription.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colDescription.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colDescription.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colDescription.AppearanceHeader.ForeColor = System.Drawing.Color.Black;
			this.colDescription.AppearanceHeader.Options.UseBackColor = true;
			this.colDescription.AppearanceHeader.Options.UseFont = true;
			this.colDescription.AppearanceHeader.Options.UseForeColor = true;
			this.colDescription.AppearanceHeader.Options.UseTextOptions = true;
			this.colDescription.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colDescription.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colDescription.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colDescription.Caption = "Description";
			this.colDescription.FieldName = "Description";
			this.colDescription.MinWidth = 25;
			this.colDescription.Name = "colDescription";
			this.colDescription.OptionsColumn.AllowEdit = false;
			this.colDescription.Visible = true;
			this.colDescription.VisibleIndex = 5;
			this.colDescription.Width = 113;
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
			this.colQty.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colQty.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colQty.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colQty.AppearanceHeader.ForeColor = System.Drawing.Color.Black;
			this.colQty.AppearanceHeader.Options.UseBackColor = true;
			this.colQty.AppearanceHeader.Options.UseFont = true;
			this.colQty.AppearanceHeader.Options.UseForeColor = true;
			this.colQty.AppearanceHeader.Options.UseTextOptions = true;
			this.colQty.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colQty.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colQty.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colQty.Caption = "Số lượng";
			this.colQty.FieldName = "Qty";
			this.colQty.MinWidth = 25;
			this.colQty.Name = "colQty";
			this.colQty.OptionsColumn.AllowEdit = false;
			this.colQty.Visible = true;
			this.colQty.VisibleIndex = 6;
			this.colQty.Width = 122;
			// 
			// colShelf
			// 
			this.colShelf.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.colShelf.AppearanceCell.Options.UseFont = true;
			this.colShelf.AppearanceCell.Options.UseTextOptions = true;
			this.colShelf.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colShelf.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colShelf.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colShelf.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colShelf.AppearanceHeader.ForeColor = System.Drawing.Color.Black;
			this.colShelf.AppearanceHeader.Options.UseBackColor = true;
			this.colShelf.AppearanceHeader.Options.UseFont = true;
			this.colShelf.AppearanceHeader.Options.UseForeColor = true;
			this.colShelf.AppearanceHeader.Options.UseTextOptions = true;
			this.colShelf.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colShelf.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colShelf.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colShelf.Caption = "Shelf";
			this.colShelf.FieldName = "Shelf";
			this.colShelf.MinWidth = 25;
			this.colShelf.Name = "colShelf";
			this.colShelf.Visible = true;
			this.colShelf.VisibleIndex = 7;
			// 
			// colLocation
			// 
			this.colLocation.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.colLocation.AppearanceCell.Options.UseFont = true;
			this.colLocation.AppearanceCell.Options.UseTextOptions = true;
			this.colLocation.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colLocation.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colLocation.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colLocation.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colLocation.AppearanceHeader.ForeColor = System.Drawing.Color.Black;
			this.colLocation.AppearanceHeader.Options.UseBackColor = true;
			this.colLocation.AppearanceHeader.Options.UseFont = true;
			this.colLocation.AppearanceHeader.Options.UseForeColor = true;
			this.colLocation.AppearanceHeader.Options.UseTextOptions = true;
			this.colLocation.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colLocation.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colLocation.Caption = "Location";
			this.colLocation.FieldName = "Location";
			this.colLocation.MinWidth = 25;
			this.colLocation.Name = "colLocation";
			this.colLocation.Visible = true;
			this.colLocation.VisibleIndex = 8;
			this.colLocation.Width = 77;
			// 
			// colOrderCode
			// 
			this.colOrderCode.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.colOrderCode.AppearanceCell.Options.UseFont = true;
			this.colOrderCode.AppearanceCell.Options.UseTextOptions = true;
			this.colOrderCode.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colOrderCode.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colOrderCode.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colOrderCode.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colOrderCode.AppearanceHeader.ForeColor = System.Drawing.Color.Black;
			this.colOrderCode.AppearanceHeader.Options.UseBackColor = true;
			this.colOrderCode.AppearanceHeader.Options.UseFont = true;
			this.colOrderCode.AppearanceHeader.Options.UseForeColor = true;
			this.colOrderCode.AppearanceHeader.Options.UseTextOptions = true;
			this.colOrderCode.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colOrderCode.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colOrderCode.Caption = "MÃ ORDER";
			this.colOrderCode.FieldName = "OrderCode";
			this.colOrderCode.MinWidth = 25;
			this.colOrderCode.Name = "colOrderCode";
			this.colOrderCode.Visible = true;
			this.colOrderCode.VisibleIndex = 9;
			this.colOrderCode.Width = 173;
			// 
			// colLot
			// 
			this.colLot.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.colLot.AppearanceCell.Options.UseFont = true;
			this.colLot.AppearanceCell.Options.UseTextOptions = true;
			this.colLot.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colLot.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colLot.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colLot.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colLot.AppearanceHeader.ForeColor = System.Drawing.Color.Black;
			this.colLot.AppearanceHeader.Options.UseBackColor = true;
			this.colLot.AppearanceHeader.Options.UseFont = true;
			this.colLot.AppearanceHeader.Options.UseForeColor = true;
			this.colLot.AppearanceHeader.Options.UseTextOptions = true;
			this.colLot.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colLot.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colLot.Caption = "Lot";
			this.colLot.FieldName = "Lot";
			this.colLot.MinWidth = 25;
			this.colLot.Name = "colLot";
			this.colLot.Visible = true;
			this.colLot.VisibleIndex = 10;
			this.colLot.Width = 161;
			// 
			// colCreateAt
			// 
			this.colCreateAt.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.colCreateAt.AppearanceCell.Options.UseFont = true;
			this.colCreateAt.AppearanceCell.Options.UseTextOptions = true;
			this.colCreateAt.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colCreateAt.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colCreateAt.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colCreateAt.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colCreateAt.AppearanceHeader.ForeColor = System.Drawing.Color.Black;
			this.colCreateAt.AppearanceHeader.Options.UseBackColor = true;
			this.colCreateAt.AppearanceHeader.Options.UseFont = true;
			this.colCreateAt.AppearanceHeader.Options.UseForeColor = true;
			this.colCreateAt.AppearanceHeader.Options.UseTextOptions = true;
			this.colCreateAt.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colCreateAt.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colCreateAt.Caption = "Ngày tạo";
			this.colCreateAt.FieldName = "CreateAt";
			this.colCreateAt.MinWidth = 25;
			this.colCreateAt.Name = "colCreateAt";
			this.colCreateAt.Visible = true;
			this.colCreateAt.VisibleIndex = 11;
			// 
			// colUserr
			// 
			this.colUserr.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.colUserr.AppearanceCell.Options.UseFont = true;
			this.colUserr.AppearanceCell.Options.UseTextOptions = true;
			this.colUserr.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colUserr.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colUserr.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colUserr.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colUserr.AppearanceHeader.ForeColor = System.Drawing.Color.Black;
			this.colUserr.AppearanceHeader.Options.UseBackColor = true;
			this.colUserr.AppearanceHeader.Options.UseFont = true;
			this.colUserr.AppearanceHeader.Options.UseForeColor = true;
			this.colUserr.AppearanceHeader.Options.UseTextOptions = true;
			this.colUserr.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colUserr.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colUserr.Caption = "Người tạo";
			this.colUserr.FieldName = "Userr";
			this.colUserr.MinWidth = 25;
			this.colUserr.Name = "colUserr";
			this.colUserr.Visible = true;
			this.colUserr.VisibleIndex = 12;
			this.colUserr.Width = 118;
			// 
			// colPID
			// 
			this.colPID.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.colPID.AppearanceCell.Options.UseFont = true;
			this.colPID.AppearanceCell.Options.UseTextOptions = true;
			this.colPID.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colPID.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colPID.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colPID.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colPID.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colPID.AppearanceHeader.Options.UseBackColor = true;
			this.colPID.AppearanceHeader.Options.UseFont = true;
			this.colPID.AppearanceHeader.Options.UseTextOptions = true;
			this.colPID.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colPID.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colPID.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colPID.Caption = "PID";
			this.colPID.FieldName = "ProductCode";
			this.colPID.Name = "colPID";
			this.colPID.Visible = true;
			this.colPID.VisibleIndex = 1;
			this.colPID.Width = 156;
			// 
			// colOrderQty
			// 
			this.colOrderQty.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.colOrderQty.AppearanceCell.Options.UseFont = true;
			this.colOrderQty.AppearanceCell.Options.UseTextOptions = true;
			this.colOrderQty.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colOrderQty.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colOrderQty.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colOrderQty.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colOrderQty.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colOrderQty.AppearanceHeader.Options.UseBackColor = true;
			this.colOrderQty.AppearanceHeader.Options.UseFont = true;
			this.colOrderQty.AppearanceHeader.Options.UseTextOptions = true;
			this.colOrderQty.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colOrderQty.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colOrderQty.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colOrderQty.Caption = "Order Q\'ty";
			this.colOrderQty.FieldName = "OrderQty";
			this.colOrderQty.Name = "colOrderQty";
			this.colOrderQty.Visible = true;
			this.colOrderQty.VisibleIndex = 2;
			this.colOrderQty.Width = 86;
			// 
			// colProductGroupCode
			// 
			this.colProductGroupCode.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.colProductGroupCode.AppearanceCell.Options.UseFont = true;
			this.colProductGroupCode.AppearanceCell.Options.UseTextOptions = true;
			this.colProductGroupCode.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colProductGroupCode.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colProductGroupCode.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colProductGroupCode.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colProductGroupCode.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colProductGroupCode.AppearanceHeader.Options.UseBackColor = true;
			this.colProductGroupCode.AppearanceHeader.Options.UseFont = true;
			this.colProductGroupCode.AppearanceHeader.Options.UseTextOptions = true;
			this.colProductGroupCode.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colProductGroupCode.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colProductGroupCode.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colProductGroupCode.Caption = "Nhóm sản phẩm";
			this.colProductGroupCode.FieldName = "ProductGroupCode";
			this.colProductGroupCode.Name = "colProductGroupCode";
			this.colProductGroupCode.Visible = true;
			this.colProductGroupCode.VisibleIndex = 13;
			// 
			// panel2
			// 
			this.panel2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.panel2.Controls.Add(this.btnSearch);
			this.panel2.Controls.Add(this.txbSearch);
			this.panel2.Controls.Add(this.dtpFrom);
			this.panel2.Controls.Add(this.dtpTo);
			this.panel2.Controls.Add(this.label1);
			this.panel2.Controls.Add(this.label3);
			this.panel2.Location = new System.Drawing.Point(0, 43);
			this.panel2.Name = "panel2";
			this.panel2.Size = new System.Drawing.Size(628, 34);
			this.panel2.TabIndex = 30;
			// 
			// btnSearch
			// 
			this.btnSearch.Appearance.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnSearch.Appearance.Options.UseFont = true;
			this.btnSearch.Location = new System.Drawing.Point(547, 5);
			this.btnSearch.Name = "btnSearch";
			this.btnSearch.Size = new System.Drawing.Size(75, 23);
			this.btnSearch.TabIndex = 25;
			this.btnSearch.Text = "Tìm kiếm";
			this.btnSearch.Click += new System.EventHandler(this.btnSearch_Click);
			// 
			// txbSearch
			// 
			this.txbSearch.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txbSearch.Location = new System.Drawing.Point(394, 6);
			this.txbSearch.Name = "txbSearch";
			this.txbSearch.Size = new System.Drawing.Size(147, 24);
			this.txbSearch.TabIndex = 24;
			this.txbSearch.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txbSearch_KeyDown);
			// 
			// dtpFrom
			// 
			this.dtpFrom.CustomFormat = "dd/MM/yyyy";
			this.dtpFrom.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.dtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
			this.dtpFrom.Location = new System.Drawing.Point(77, 6);
			this.dtpFrom.Name = "dtpFrom";
			this.dtpFrom.Size = new System.Drawing.Size(106, 24);
			this.dtpFrom.TabIndex = 36;
			this.dtpFrom.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtpFrom_KeyPress);
			// 
			// dtpTo
			// 
			this.dtpTo.CustomFormat = "dd/MM/yyyy";
			this.dtpTo.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.dtpTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
			this.dtpTo.Location = new System.Drawing.Point(274, 6);
			this.dtpTo.Name = "dtpTo";
			this.dtpTo.Size = new System.Drawing.Size(99, 24);
			this.dtpTo.TabIndex = 35;
			this.dtpTo.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtpTo_KeyPress);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.label1.Location = new System.Drawing.Point(205, 8);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(78, 18);
			this.label1.TabIndex = 34;
			this.label1.Text = "Đến ngày: ";
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.label3.Location = new System.Drawing.Point(12, 8);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(68, 18);
			this.label3.TabIndex = 33;
			this.label3.Text = "Từ ngày: ";
			// 
			// chkExportExcelAltax
			// 
			this.chkExportExcelAltax.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.chkExportExcelAltax.AutoSize = true;
			this.chkExportExcelAltax.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.chkExportExcelAltax.Location = new System.Drawing.Point(831, 54);
			this.chkExportExcelAltax.Name = "chkExportExcelAltax";
			this.chkExportExcelAltax.Size = new System.Drawing.Size(111, 17);
			this.chkExportExcelAltax.TabIndex = 45;
			this.chkExportExcelAltax.Text = "Export Excel Altax";
			this.chkExportExcelAltax.UseVisualStyleBackColor = true;
			this.chkExportExcelAltax.CheckedChanged += new System.EventHandler(this.chkExportExcel_CheckedChanged);
			// 
			// chkExportExcelHyp
			// 
			this.chkExportExcelHyp.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.chkExportExcelHyp.AutoSize = true;
			this.chkExportExcelHyp.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.chkExportExcelHyp.Location = new System.Drawing.Point(706, 54);
			this.chkExportExcelHyp.Name = "chkExportExcelHyp";
			this.chkExportExcelHyp.Size = new System.Drawing.Size(107, 17);
			this.chkExportExcelHyp.TabIndex = 46;
			this.chkExportExcelHyp.Text = "Export Excel Hyp";
			this.chkExportExcelHyp.UseVisualStyleBackColor = true;
			this.chkExportExcelHyp.CheckedChanged += new System.EventHandler(this.chkExportExcelHyp_CheckedChanged);
			// 
			// btnDeleteDuplicate
			// 
			this.btnDeleteDuplicate.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnDeleteDuplicate.Appearance.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnDeleteDuplicate.Appearance.Options.UseFont = true;
			this.btnDeleteDuplicate.Location = new System.Drawing.Point(810, 9);
			this.btnDeleteDuplicate.Name = "btnDeleteDuplicate";
			this.btnDeleteDuplicate.Size = new System.Drawing.Size(132, 27);
			this.btnDeleteDuplicate.TabIndex = 37;
			this.btnDeleteDuplicate.Text = "Xóa trùng nhau";
			this.btnDeleteDuplicate.Click += new System.EventHandler(this.btnDeleteDuplicate_Click);
			// 
			// frmAssemblyOrder
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(954, 450);
			this.Controls.Add(this.btnDeleteDuplicate);
			this.Controls.Add(this.chkExportExcelHyp);
			this.Controls.Add(this.chkExportExcelAltax);
			this.Controls.Add(this.panel2);
			this.Controls.Add(this.grdData);
			this.Controls.Add(this.toolStrip1);
			this.Margin = new System.Windows.Forms.Padding(4);
			this.Name = "frmAssemblyOrder";
			this.Text = "DANH SÁCH LẮP RÁP THEO ORDER";
			this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
			this.Load += new System.EventHandler(this.frmAssemblyOrder_Load);
			this.toolStrip1.ResumeLayout(false);
			this.toolStrip1.PerformLayout();
			((System.ComponentModel.ISupportInitialize)(this.grdData)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grvData)).EndInit();
			this.panel2.ResumeLayout(false);
			this.panel2.PerformLayout();
			this.ResumeLayout(false);
			this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ToolStrip toolStrip1;
        private System.Windows.Forms.ToolStripButton btnCreatePlan;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripButton btnEditPlan;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator7;
        private System.Windows.Forms.ToolStripButton btnDelPlan;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator8;
        private System.Windows.Forms.ToolStripButton btnExportExcel;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator11;
        private System.Windows.Forms.ToolStripButton btnImportExcel;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator12;
        private DevExpress.XtraGrid.GridControl grdData;
        private DevExpress.XtraGrid.Views.Grid.GridView grvData;
        private DevExpress.XtraGrid.Columns.GridColumn colID;
        private DevExpress.XtraGrid.Columns.GridColumn colAssemblyOrder;
        private DevExpress.XtraGrid.Columns.GridColumn colArticleID;
        private DevExpress.XtraGrid.Columns.GridColumn colCnt;
        private DevExpress.XtraGrid.Columns.GridColumn colDescription;
        private DevExpress.XtraGrid.Columns.GridColumn colQty;
		private DevExpress.XtraGrid.Columns.GridColumn colShelf;
		private DevExpress.XtraGrid.Columns.GridColumn colLocation;
		private DevExpress.XtraGrid.Columns.GridColumn colOrderCode;
		private DevExpress.XtraGrid.Columns.GridColumn colLot;
		private DevExpress.XtraGrid.Columns.GridColumn colCreateAt;
		private DevExpress.XtraGrid.Columns.GridColumn colUserr;
		private System.Windows.Forms.Timer timer1;
		private System.Windows.Forms.Panel panel2;
		private DevExpress.XtraEditors.SimpleButton btnSearch;
		private System.Windows.Forms.TextBox txbSearch;
		private System.Windows.Forms.DateTimePicker dtpFrom;
		private System.Windows.Forms.DateTimePicker dtpTo;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.CheckBox chkExportExcelAltax;
		private DevExpress.XtraGrid.Columns.GridColumn colPID;
		private DevExpress.XtraGrid.Columns.GridColumn colOrderQty;
		private System.Windows.Forms.CheckBox chkExportExcelHyp;
		private DevExpress.XtraGrid.Columns.GridColumn colProductGroupCode;
		private DevExpress.XtraEditors.SimpleButton btnDeleteDuplicate;
	}
}