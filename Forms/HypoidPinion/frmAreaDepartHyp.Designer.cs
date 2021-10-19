
namespace BMS
{
	partial class frmAreaDepartHyp
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
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmAreaDepartHyp));
			this.toolStrip1 = new System.Windows.Forms.ToolStrip();
			this.btnCreateDep = new System.Windows.Forms.ToolStripButton();
			this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
			this.btnEditDep = new System.Windows.Forms.ToolStripButton();
			this.toolStripSeparator7 = new System.Windows.Forms.ToolStripSeparator();
			this.btnDelDep = new System.Windows.Forms.ToolStripButton();
			this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
			this.btnRefresh = new System.Windows.Forms.ToolStripButton();
			this.toolStripSeparator9 = new System.Windows.Forms.ToolStripSeparator();
			this.btnCreateArea = new System.Windows.Forms.ToolStripButton();
			this.toolStripSeparator6 = new System.Windows.Forms.ToolStripSeparator();
			this.btnEditArea = new System.Windows.Forms.ToolStripButton();
			this.toolStripSeparator8 = new System.Windows.Forms.ToolStripSeparator();
			this.btnDelArea = new System.Windows.Forms.ToolStripButton();
			this.gridView1 = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.splitContainer1 = new System.Windows.Forms.SplitContainer();
			this.dtgvArea = new DevExpress.XtraGrid.GridControl();
			this.gvArea = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.colIDAr = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colAreaCodeAr = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colManagerIDAr = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colManagerCodeAr = new DevExpress.XtraGrid.Columns.GridColumn();
			this.gridColumn2 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colAreaName = new DevExpress.XtraGrid.Columns.GridColumn();
			this.labelControl1 = new DevExpress.XtraEditors.LabelControl();
			this.dtgvDepartment = new DevExpress.XtraGrid.GridControl();
			this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
			this.btnCheckStage = new System.Windows.Forms.ToolStripMenuItem();
			this.gvDepartment = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.colIDDep = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colDepartmentCodeDep = new DevExpress.XtraGrid.Columns.GridColumn();
			this.gridColumn3 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.gridColumn4 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colSTT = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colAreaIDDep = new DevExpress.XtraGrid.Columns.GridColumn();
			this.gridColumn5 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colDepartmentName = new DevExpress.XtraGrid.Columns.GridColumn();
			this.gridColumn1 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.cbShowAll = new System.Windows.Forms.CheckBox();
			this.panel6 = new System.Windows.Forms.Panel();
			this.btnCancelSearch = new DevExpress.XtraEditors.SimpleButton();
			this.label1 = new System.Windows.Forms.Label();
			this.txbSearch = new System.Windows.Forms.TextBox();
			this.btnSearch = new DevExpress.XtraEditors.SimpleButton();
			this.toolStrip1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.gridView1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
			this.splitContainer1.Panel1.SuspendLayout();
			this.splitContainer1.Panel2.SuspendLayout();
			this.splitContainer1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.dtgvArea)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gvArea)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtgvDepartment)).BeginInit();
			this.contextMenuStrip1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.gvDepartment)).BeginInit();
			this.panel6.SuspendLayout();
			this.SuspendLayout();
			// 
			// toolStrip1
			// 
			this.toolStrip1.AutoSize = false;
			this.toolStrip1.BackColor = System.Drawing.Color.WhiteSmoke;
			this.toolStrip1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.toolStrip1.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden;
			this.toolStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
			this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.btnCreateDep,
            this.toolStripSeparator1,
            this.btnEditDep,
            this.toolStripSeparator7,
            this.btnDelDep,
            this.toolStripSeparator2,
            this.btnRefresh,
            this.toolStripSeparator9,
            this.btnCreateArea,
            this.toolStripSeparator6,
            this.btnEditArea,
            this.toolStripSeparator8,
            this.btnDelArea});
			this.toolStrip1.LayoutStyle = System.Windows.Forms.ToolStripLayoutStyle.Flow;
			this.toolStrip1.Location = new System.Drawing.Point(0, 0);
			this.toolStrip1.Name = "toolStrip1";
			this.toolStrip1.Padding = new System.Windows.Forms.Padding(0, 0, 2, 0);
			this.toolStrip1.RenderMode = System.Windows.Forms.ToolStripRenderMode.Professional;
			this.toolStrip1.Size = new System.Drawing.Size(1028, 42);
			this.toolStrip1.TabIndex = 30;
			this.toolStrip1.Text = "toolStrip2";
			// 
			// btnCreateDep
			// 
			this.btnCreateDep.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnCreateDep.ForeColor = System.Drawing.Color.Black;
			this.btnCreateDep.Image = ((System.Drawing.Image)(resources.GetObject("btnCreateDep.Image")));
			this.btnCreateDep.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnCreateDep.Name = "btnCreateDep";
			this.btnCreateDep.Size = new System.Drawing.Size(104, 41);
			this.btnCreateDep.Tag = "frmProduct_AddProductH";
			this.btnCreateDep.Text = "Tạo bộ phận";
			this.btnCreateDep.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnCreateDep.Click += new System.EventHandler(this.btnCreateDep_Click);
			// 
			// toolStripSeparator1
			// 
			this.toolStripSeparator1.Name = "toolStripSeparator1";
			this.toolStripSeparator1.Size = new System.Drawing.Size(6, 23);
			// 
			// btnEditDep
			// 
			this.btnEditDep.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnEditDep.ForeColor = System.Drawing.Color.Black;
			this.btnEditDep.Image = ((System.Drawing.Image)(resources.GetObject("btnEditDep.Image")));
			this.btnEditDep.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnEditDep.Name = "btnEditDep";
			this.btnEditDep.Size = new System.Drawing.Size(104, 41);
			this.btnEditDep.Tag = "frmProduct_EditProductH";
			this.btnEditDep.Text = "Sửa bộ phận";
			this.btnEditDep.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnEditDep.Click += new System.EventHandler(this.btnEditDep_Click);
			// 
			// toolStripSeparator7
			// 
			this.toolStripSeparator7.Name = "toolStripSeparator7";
			this.toolStripSeparator7.Size = new System.Drawing.Size(6, 23);
			// 
			// btnDelDep
			// 
			this.btnDelDep.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnDelDep.ForeColor = System.Drawing.Color.Black;
			this.btnDelDep.Image = ((System.Drawing.Image)(resources.GetObject("btnDelDep.Image")));
			this.btnDelDep.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnDelDep.Name = "btnDelDep";
			this.btnDelDep.Size = new System.Drawing.Size(104, 41);
			this.btnDelDep.Tag = "frmProduct_DeleteProductH";
			this.btnDelDep.Text = "Xóa bộ phận";
			this.btnDelDep.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnDelDep.Click += new System.EventHandler(this.btnDelDep_Click);
			// 
			// toolStripSeparator2
			// 
			this.toolStripSeparator2.AutoSize = false;
			this.toolStripSeparator2.Name = "toolStripSeparator2";
			this.toolStripSeparator2.Size = new System.Drawing.Size(6, 40);
			// 
			// btnRefresh
			// 
			this.btnRefresh.AutoSize = false;
			this.btnRefresh.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Bold);
			this.btnRefresh.Image = ((System.Drawing.Image)(resources.GetObject("btnRefresh.Image")));
			this.btnRefresh.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnRefresh.Name = "btnRefresh";
			this.btnRefresh.Size = new System.Drawing.Size(91, 44);
			this.btnRefresh.Tag = "";
			this.btnRefresh.Text = "Làm mới";
			this.btnRefresh.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnRefresh.Click += new System.EventHandler(this.btnRefresh_Click);
			// 
			// toolStripSeparator9
			// 
			this.toolStripSeparator9.AutoSize = false;
			this.toolStripSeparator9.Name = "toolStripSeparator9";
			this.toolStripSeparator9.Size = new System.Drawing.Size(6, 40);
			// 
			// btnCreateArea
			// 
			this.btnCreateArea.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnCreateArea.ForeColor = System.Drawing.Color.Black;
			this.btnCreateArea.Image = ((System.Drawing.Image)(resources.GetObject("btnCreateArea.Image")));
			this.btnCreateArea.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnCreateArea.Name = "btnCreateArea";
			this.btnCreateArea.Size = new System.Drawing.Size(101, 41);
			this.btnCreateArea.Tag = "frmProduct_AddProductH";
			this.btnCreateArea.Text = "Tạo khu vực";
			this.btnCreateArea.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnCreateArea.Click += new System.EventHandler(this.btnCreateArea_Click);
			// 
			// toolStripSeparator6
			// 
			this.toolStripSeparator6.Name = "toolStripSeparator6";
			this.toolStripSeparator6.Size = new System.Drawing.Size(6, 23);
			// 
			// btnEditArea
			// 
			this.btnEditArea.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnEditArea.ForeColor = System.Drawing.Color.Black;
			this.btnEditArea.Image = ((System.Drawing.Image)(resources.GetObject("btnEditArea.Image")));
			this.btnEditArea.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnEditArea.Name = "btnEditArea";
			this.btnEditArea.Size = new System.Drawing.Size(101, 41);
			this.btnEditArea.Tag = "frmProduct_EditProductH";
			this.btnEditArea.Text = "Sửa khu vực";
			this.btnEditArea.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnEditArea.Click += new System.EventHandler(this.btnEditArea_Click);
			// 
			// toolStripSeparator8
			// 
			this.toolStripSeparator8.Name = "toolStripSeparator8";
			this.toolStripSeparator8.Size = new System.Drawing.Size(6, 23);
			// 
			// btnDelArea
			// 
			this.btnDelArea.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnDelArea.ForeColor = System.Drawing.Color.Black;
			this.btnDelArea.Image = ((System.Drawing.Image)(resources.GetObject("btnDelArea.Image")));
			this.btnDelArea.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnDelArea.Name = "btnDelArea";
			this.btnDelArea.Size = new System.Drawing.Size(101, 41);
			this.btnDelArea.Tag = "frmProduct_DeleteProductH";
			this.btnDelArea.Text = "Xóa khu vực";
			this.btnDelArea.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnDelArea.Click += new System.EventHandler(this.btnDelArea_Click);
			// 
			// gridView1
			// 
			this.gridView1.Name = "gridView1";
			// 
			// splitContainer1
			// 
			this.splitContainer1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.splitContainer1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.splitContainer1.Location = new System.Drawing.Point(0, 44);
			this.splitContainer1.Margin = new System.Windows.Forms.Padding(2);
			this.splitContainer1.Name = "splitContainer1";
			// 
			// splitContainer1.Panel1
			// 
			this.splitContainer1.Panel1.BackColor = System.Drawing.Color.WhiteSmoke;
			this.splitContainer1.Panel1.Controls.Add(this.dtgvArea);
			this.splitContainer1.Panel1.Controls.Add(this.labelControl1);
			this.splitContainer1.Panel1MinSize = 15;
			// 
			// splitContainer1.Panel2
			// 
			this.splitContainer1.Panel2.Controls.Add(this.dtgvDepartment);
			this.splitContainer1.Panel2.Controls.Add(this.cbShowAll);
			this.splitContainer1.Panel2.Controls.Add(this.panel6);
			this.splitContainer1.Size = new System.Drawing.Size(1052, 556);
			this.splitContainer1.SplitterDistance = 245;
			this.splitContainer1.SplitterWidth = 3;
			this.splitContainer1.TabIndex = 36;
			// 
			// dtgvArea
			// 
			this.dtgvArea.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.dtgvArea.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.dtgvArea.Location = new System.Drawing.Point(3, 45);
			this.dtgvArea.LookAndFeel.Style = DevExpress.LookAndFeel.LookAndFeelStyle.Flat;
			this.dtgvArea.MainView = this.gvArea;
			this.dtgvArea.Name = "dtgvArea";
			this.dtgvArea.Size = new System.Drawing.Size(241, 506);
			this.dtgvArea.TabIndex = 36;
			this.dtgvArea.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.gvArea});
			this.dtgvArea.DoubleClick += new System.EventHandler(this.dtgvArea_DoubleClick);
			// 
			// gvArea
			// 
			this.gvArea.ColumnPanelRowHeight = 40;
			this.gvArea.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
            this.colIDAr,
            this.colAreaCodeAr,
            this.colManagerIDAr,
            this.colManagerCodeAr,
            this.gridColumn2,
            this.colAreaName});
			this.gvArea.GridControl = this.dtgvArea;
			this.gvArea.Name = "gvArea";
			this.gvArea.OptionsBehavior.AllowAddRows = DevExpress.Utils.DefaultBoolean.False;
			this.gvArea.OptionsBehavior.AllowDeleteRows = DevExpress.Utils.DefaultBoolean.False;
			this.gvArea.OptionsBehavior.AllowFixedGroups = DevExpress.Utils.DefaultBoolean.False;
			this.gvArea.OptionsBehavior.AllowIncrementalSearch = true;
			this.gvArea.OptionsBehavior.Editable = false;
			this.gvArea.OptionsCustomization.AllowColumnMoving = false;
			this.gvArea.OptionsCustomization.AllowQuickHideColumns = false;
			this.gvArea.OptionsFind.AlwaysVisible = true;
			this.gvArea.OptionsFind.ShowCloseButton = false;
			this.gvArea.OptionsView.ShowGroupPanel = false;
			this.gvArea.RowHeight = 25;
			this.gvArea.ScrollStyle = DevExpress.XtraGrid.Views.Grid.ScrollStyleFlags.LiveVertScroll;
			this.gvArea.FocusedRowChanged += new DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventHandler(this.gvArea_FocusedRowChanged);
			// 
			// colIDAr
			// 
			this.colIDAr.Caption = "ID";
			this.colIDAr.FieldName = "ID";
			this.colIDAr.Name = "colIDAr";
			this.colIDAr.OptionsColumn.AllowEdit = false;
			// 
			// colAreaCodeAr
			// 
			this.colAreaCodeAr.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.colAreaCodeAr.AppearanceCell.Options.UseFont = true;
			this.colAreaCodeAr.AppearanceCell.Options.UseTextOptions = true;
			this.colAreaCodeAr.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colAreaCodeAr.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colAreaCodeAr.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colAreaCodeAr.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.colAreaCodeAr.AppearanceHeader.ForeColor = System.Drawing.Color.Black;
			this.colAreaCodeAr.AppearanceHeader.Options.UseFont = true;
			this.colAreaCodeAr.AppearanceHeader.Options.UseForeColor = true;
			this.colAreaCodeAr.AppearanceHeader.Options.UseTextOptions = true;
			this.colAreaCodeAr.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colAreaCodeAr.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colAreaCodeAr.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colAreaCodeAr.Caption = "MÃ KHU VỰC";
			this.colAreaCodeAr.FieldName = "AreaCode";
			this.colAreaCodeAr.Name = "colAreaCodeAr";
			this.colAreaCodeAr.OptionsColumn.AllowEdit = false;
			this.colAreaCodeAr.Visible = true;
			this.colAreaCodeAr.VisibleIndex = 0;
			this.colAreaCodeAr.Width = 119;
			// 
			// colManagerIDAr
			// 
			this.colManagerIDAr.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.colManagerIDAr.AppearanceCell.Options.UseFont = true;
			this.colManagerIDAr.AppearanceCell.Options.UseTextOptions = true;
			this.colManagerIDAr.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colManagerIDAr.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colManagerIDAr.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colManagerIDAr.AppearanceHeader.ForeColor = System.Drawing.Color.Black;
			this.colManagerIDAr.AppearanceHeader.Options.UseFont = true;
			this.colManagerIDAr.AppearanceHeader.Options.UseForeColor = true;
			this.colManagerIDAr.AppearanceHeader.Options.UseTextOptions = true;
			this.colManagerIDAr.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colManagerIDAr.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colManagerIDAr.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colManagerIDAr.Caption = "Mô tả";
			this.colManagerIDAr.FieldName = "ManagerID";
			this.colManagerIDAr.Name = "colManagerIDAr";
			this.colManagerIDAr.OptionsColumn.AllowEdit = false;
			this.colManagerIDAr.Width = 338;
			// 
			// colManagerCodeAr
			// 
			this.colManagerCodeAr.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.colManagerCodeAr.AppearanceCell.Options.UseFont = true;
			this.colManagerCodeAr.AppearanceCell.Options.UseTextOptions = true;
			this.colManagerCodeAr.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colManagerCodeAr.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colManagerCodeAr.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colManagerCodeAr.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.colManagerCodeAr.AppearanceHeader.ForeColor = System.Drawing.Color.Black;
			this.colManagerCodeAr.AppearanceHeader.Options.UseFont = true;
			this.colManagerCodeAr.AppearanceHeader.Options.UseForeColor = true;
			this.colManagerCodeAr.AppearanceHeader.Options.UseTextOptions = true;
			this.colManagerCodeAr.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colManagerCodeAr.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colManagerCodeAr.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colManagerCodeAr.Caption = "NGƯỜI PHỤ TRÁCH";
			this.colManagerCodeAr.FieldName = "ManagerCode";
			this.colManagerCodeAr.Name = "colManagerCodeAr";
			this.colManagerCodeAr.OptionsColumn.AllowEdit = false;
			this.colManagerCodeAr.Visible = true;
			this.colManagerCodeAr.VisibleIndex = 2;
			this.colManagerCodeAr.Width = 145;
			// 
			// gridColumn2
			// 
			this.gridColumn2.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.gridColumn2.AppearanceCell.Options.UseFont = true;
			this.gridColumn2.AppearanceCell.Options.UseTextOptions = true;
			this.gridColumn2.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.gridColumn2.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.gridColumn2.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.gridColumn2.AppearanceHeader.ForeColor = System.Drawing.Color.Black;
			this.gridColumn2.AppearanceHeader.Options.UseFont = true;
			this.gridColumn2.AppearanceHeader.Options.UseForeColor = true;
			this.gridColumn2.AppearanceHeader.Options.UseTextOptions = true;
			this.gridColumn2.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.gridColumn2.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.gridColumn2.Caption = "STT";
			this.gridColumn2.FieldName = "STT";
			this.gridColumn2.Name = "gridColumn2";
			this.gridColumn2.Width = 41;
			// 
			// colAreaName
			// 
			this.colAreaName.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.colAreaName.AppearanceCell.Options.UseFont = true;
			this.colAreaName.AppearanceCell.Options.UseTextOptions = true;
			this.colAreaName.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colAreaName.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colAreaName.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colAreaName.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 10.2F, System.Drawing.FontStyle.Bold);
			this.colAreaName.AppearanceHeader.ForeColor = System.Drawing.Color.Black;
			this.colAreaName.AppearanceHeader.Options.UseFont = true;
			this.colAreaName.AppearanceHeader.Options.UseForeColor = true;
			this.colAreaName.AppearanceHeader.Options.UseTextOptions = true;
			this.colAreaName.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colAreaName.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colAreaName.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colAreaName.Caption = "TÊN KHU VỰC";
			this.colAreaName.FieldName = "AreaName";
			this.colAreaName.Name = "colAreaName";
			this.colAreaName.Visible = true;
			this.colAreaName.VisibleIndex = 1;
			// 
			// labelControl1
			// 
			this.labelControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)));
			this.labelControl1.Appearance.Font = new System.Drawing.Font("Tahoma", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.labelControl1.Location = new System.Drawing.Point(60, 13);
			this.labelControl1.Name = "labelControl1";
			this.labelControl1.Size = new System.Drawing.Size(130, 17);
			this.labelControl1.TabIndex = 19;
			this.labelControl1.Text = "Danh sách khu vực";
			// 
			// dtgvDepartment
			// 
			this.dtgvDepartment.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.dtgvDepartment.ContextMenuStrip = this.contextMenuStrip1;
			this.dtgvDepartment.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.dtgvDepartment.Location = new System.Drawing.Point(3, 45);
			this.dtgvDepartment.LookAndFeel.Style = DevExpress.LookAndFeel.LookAndFeelStyle.Flat;
			this.dtgvDepartment.MainView = this.gvDepartment;
			this.dtgvDepartment.Name = "dtgvDepartment";
			this.dtgvDepartment.Size = new System.Drawing.Size(777, 524);
			this.dtgvDepartment.TabIndex = 37;
			this.dtgvDepartment.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.gvDepartment});
			// 
			// contextMenuStrip1
			// 
			this.contextMenuStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
			this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.btnCheckStage});
			this.contextMenuStrip1.Name = "contextMenuStrip1";
			this.contextMenuStrip1.Size = new System.Drawing.Size(159, 26);
			// 
			// btnCheckStage
			// 
			this.btnCheckStage.Name = "btnCheckStage";
			this.btnCheckStage.Size = new System.Drawing.Size(158, 22);
			this.btnCheckStage.Tag = "Module_Film";
			this.btnCheckStage.Text = "Xem công đoạn";
			this.btnCheckStage.Click += new System.EventHandler(this.btnCheckStage_Click);
			// 
			// gvDepartment
			// 
			this.gvDepartment.ColumnPanelRowHeight = 40;
			this.gvDepartment.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
            this.colIDDep,
            this.colDepartmentCodeDep,
            this.gridColumn3,
            this.gridColumn4,
            this.colSTT,
            this.colAreaIDDep,
            this.gridColumn5,
            this.colDepartmentName,
            this.gridColumn1});
			this.gvDepartment.GridControl = this.dtgvDepartment;
			this.gvDepartment.Name = "gvDepartment";
			this.gvDepartment.OptionsBehavior.AllowAddRows = DevExpress.Utils.DefaultBoolean.False;
			this.gvDepartment.OptionsBehavior.AllowDeleteRows = DevExpress.Utils.DefaultBoolean.False;
			this.gvDepartment.OptionsBehavior.AllowFixedGroups = DevExpress.Utils.DefaultBoolean.False;
			this.gvDepartment.OptionsBehavior.AllowIncrementalSearch = true;
			this.gvDepartment.OptionsBehavior.Editable = false;
			this.gvDepartment.OptionsCustomization.AllowColumnMoving = false;
			this.gvDepartment.OptionsCustomization.AllowQuickHideColumns = false;
			this.gvDepartment.OptionsFind.AlwaysVisible = true;
			this.gvDepartment.OptionsFind.ShowCloseButton = false;
			this.gvDepartment.OptionsView.ShowGroupPanel = false;
			this.gvDepartment.RowHeight = 25;
			this.gvDepartment.ScrollStyle = DevExpress.XtraGrid.Views.Grid.ScrollStyleFlags.LiveVertScroll;
			this.gvDepartment.DoubleClick += new System.EventHandler(this.gvDepartment_DoubleClick);
			// 
			// colIDDep
			// 
			this.colIDDep.Caption = "ID";
			this.colIDDep.FieldName = "ID";
			this.colIDDep.Name = "colIDDep";
			this.colIDDep.OptionsColumn.AllowEdit = false;
			// 
			// colDepartmentCodeDep
			// 
			this.colDepartmentCodeDep.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.colDepartmentCodeDep.AppearanceCell.Options.UseFont = true;
			this.colDepartmentCodeDep.AppearanceCell.Options.UseTextOptions = true;
			this.colDepartmentCodeDep.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colDepartmentCodeDep.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colDepartmentCodeDep.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colDepartmentCodeDep.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colDepartmentCodeDep.AppearanceHeader.ForeColor = System.Drawing.Color.Black;
			this.colDepartmentCodeDep.AppearanceHeader.Options.UseFont = true;
			this.colDepartmentCodeDep.AppearanceHeader.Options.UseForeColor = true;
			this.colDepartmentCodeDep.AppearanceHeader.Options.UseTextOptions = true;
			this.colDepartmentCodeDep.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colDepartmentCodeDep.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colDepartmentCodeDep.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colDepartmentCodeDep.Caption = "MÃ BỘ PHẬN";
			this.colDepartmentCodeDep.FieldName = "DepartmentCode";
			this.colDepartmentCodeDep.Name = "colDepartmentCodeDep";
			this.colDepartmentCodeDep.OptionsColumn.AllowEdit = false;
			this.colDepartmentCodeDep.Visible = true;
			this.colDepartmentCodeDep.VisibleIndex = 0;
			this.colDepartmentCodeDep.Width = 148;
			// 
			// gridColumn3
			// 
			this.gridColumn3.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.gridColumn3.AppearanceCell.Options.UseFont = true;
			this.gridColumn3.AppearanceCell.Options.UseTextOptions = true;
			this.gridColumn3.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.gridColumn3.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.gridColumn3.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.gridColumn3.AppearanceHeader.ForeColor = System.Drawing.Color.Black;
			this.gridColumn3.AppearanceHeader.Options.UseFont = true;
			this.gridColumn3.AppearanceHeader.Options.UseForeColor = true;
			this.gridColumn3.AppearanceHeader.Options.UseTextOptions = true;
			this.gridColumn3.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.gridColumn3.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.gridColumn3.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.gridColumn3.Caption = "Mô tả";
			this.gridColumn3.FieldName = "ManagerID";
			this.gridColumn3.Name = "gridColumn3";
			this.gridColumn3.OptionsColumn.AllowEdit = false;
			this.gridColumn3.Width = 338;
			// 
			// gridColumn4
			// 
			this.gridColumn4.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.gridColumn4.AppearanceCell.Options.UseFont = true;
			this.gridColumn4.AppearanceCell.Options.UseTextOptions = true;
			this.gridColumn4.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.gridColumn4.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.gridColumn4.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.gridColumn4.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.gridColumn4.AppearanceHeader.ForeColor = System.Drawing.Color.Black;
			this.gridColumn4.AppearanceHeader.Options.UseFont = true;
			this.gridColumn4.AppearanceHeader.Options.UseForeColor = true;
			this.gridColumn4.AppearanceHeader.Options.UseTextOptions = true;
			this.gridColumn4.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.gridColumn4.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.gridColumn4.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.gridColumn4.Caption = "NGƯỜI PHỤ TRÁCH";
			this.gridColumn4.FieldName = "ManagerCode";
			this.gridColumn4.Name = "gridColumn4";
			this.gridColumn4.OptionsColumn.AllowEdit = false;
			this.gridColumn4.Visible = true;
			this.gridColumn4.VisibleIndex = 3;
			this.gridColumn4.Width = 147;
			// 
			// colSTT
			// 
			this.colSTT.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.colSTT.AppearanceCell.Options.UseFont = true;
			this.colSTT.AppearanceCell.Options.UseTextOptions = true;
			this.colSTT.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colSTT.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colSTT.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colSTT.AppearanceHeader.ForeColor = System.Drawing.Color.Black;
			this.colSTT.AppearanceHeader.Options.UseFont = true;
			this.colSTT.AppearanceHeader.Options.UseForeColor = true;
			this.colSTT.AppearanceHeader.Options.UseTextOptions = true;
			this.colSTT.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colSTT.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colSTT.Caption = "STT";
			this.colSTT.FieldName = "STT";
			this.colSTT.Name = "colSTT";
			this.colSTT.Width = 151;
			// 
			// colAreaIDDep
			// 
			this.colAreaIDDep.Caption = "gridColumn1";
			this.colAreaIDDep.FieldName = "AreaID";
			this.colAreaIDDep.Name = "colAreaIDDep";
			// 
			// gridColumn5
			// 
			this.gridColumn5.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.gridColumn5.AppearanceCell.Options.UseFont = true;
			this.gridColumn5.AppearanceCell.Options.UseTextOptions = true;
			this.gridColumn5.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.gridColumn5.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.gridColumn5.AppearanceHeader.Options.UseFont = true;
			this.gridColumn5.AppearanceHeader.Options.UseTextOptions = true;
			this.gridColumn5.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.gridColumn5.Caption = "MÃ KHU VỰC";
			this.gridColumn5.FieldName = "AreaCode";
			this.gridColumn5.Name = "gridColumn5";
			this.gridColumn5.Visible = true;
			this.gridColumn5.VisibleIndex = 2;
			this.gridColumn5.Width = 148;
			// 
			// colDepartmentName
			// 
			this.colDepartmentName.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.colDepartmentName.AppearanceCell.Options.UseFont = true;
			this.colDepartmentName.AppearanceCell.Options.UseTextOptions = true;
			this.colDepartmentName.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colDepartmentName.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colDepartmentName.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colDepartmentName.AppearanceHeader.Options.UseFont = true;
			this.colDepartmentName.AppearanceHeader.Options.UseTextOptions = true;
			this.colDepartmentName.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colDepartmentName.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colDepartmentName.Caption = "TÊN BỘ PHẬN";
			this.colDepartmentName.FieldName = "DepartmentName";
			this.colDepartmentName.Name = "colDepartmentName";
			this.colDepartmentName.Visible = true;
			this.colDepartmentName.VisibleIndex = 1;
			this.colDepartmentName.Width = 156;
			// 
			// gridColumn1
			// 
			this.gridColumn1.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.gridColumn1.AppearanceCell.Options.UseFont = true;
			this.gridColumn1.AppearanceCell.Options.UseTextOptions = true;
			this.gridColumn1.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.gridColumn1.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.gridColumn1.AppearanceHeader.Options.UseFont = true;
			this.gridColumn1.AppearanceHeader.Options.UseTextOptions = true;
			this.gridColumn1.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.gridColumn1.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.gridColumn1.Caption = "MÔ TẢ";
			this.gridColumn1.FieldName = "Description";
			this.gridColumn1.Name = "gridColumn1";
			this.gridColumn1.Visible = true;
			this.gridColumn1.VisibleIndex = 4;
			this.gridColumn1.Width = 180;
			// 
			// cbShowAll
			// 
			this.cbShowAll.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.cbShowAll.AutoSize = true;
			this.cbShowAll.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.cbShowAll.Location = new System.Drawing.Point(668, 15);
			this.cbShowAll.Margin = new System.Windows.Forms.Padding(2);
			this.cbShowAll.Name = "cbShowAll";
			this.cbShowAll.Size = new System.Drawing.Size(100, 19);
			this.cbShowAll.TabIndex = 36;
			this.cbShowAll.Text = "Hiển thị tất cả";
			this.cbShowAll.UseVisualStyleBackColor = true;
			this.cbShowAll.CheckedChanged += new System.EventHandler(this.cbShowAll_CheckedChanged);
			// 
			// panel6
			// 
			this.panel6.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.panel6.Controls.Add(this.btnCancelSearch);
			this.panel6.Controls.Add(this.label1);
			this.panel6.Controls.Add(this.txbSearch);
			this.panel6.Controls.Add(this.btnSearch);
			this.panel6.Location = new System.Drawing.Point(3, 6);
			this.panel6.Name = "panel6";
			this.panel6.Size = new System.Drawing.Size(558, 36);
			this.panel6.TabIndex = 35;
			// 
			// btnCancelSearch
			// 
			this.btnCancelSearch.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelSearch.Location = new System.Drawing.Point(453, 6);
			this.btnCancelSearch.Name = "btnCancelSearch";
			this.btnCancelSearch.Size = new System.Drawing.Size(63, 20);
			this.btnCancelSearch.TabIndex = 28;
			this.btnCancelSearch.Text = "Hủy";
			this.btnCancelSearch.Click += new System.EventHandler(this.btnCancelSearch_Click);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.label1.Location = new System.Drawing.Point(3, 8);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(64, 16);
			this.label1.TabIndex = 27;
			this.label1.Text = "Từ khóa";
			// 
			// txbSearch
			// 
			this.txbSearch.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txbSearch.Location = new System.Drawing.Point(73, 5);
			this.txbSearch.Name = "txbSearch";
			this.txbSearch.Size = new System.Drawing.Size(306, 24);
			this.txbSearch.TabIndex = 26;
			this.txbSearch.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txbSearch_KeyPress);
			// 
			// btnSearch
			// 
			this.btnSearch.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSearch.Location = new System.Drawing.Point(384, 6);
			this.btnSearch.Name = "btnSearch";
			this.btnSearch.Size = new System.Drawing.Size(63, 20);
			this.btnSearch.TabIndex = 25;
			this.btnSearch.Text = "Tìm kiếm";
			this.btnSearch.Click += new System.EventHandler(this.btnSearch_Click);
			// 
			// frmAreaDepartHyp
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(1028, 602);
			this.Controls.Add(this.toolStrip1);
			this.Controls.Add(this.splitContainer1);
			this.Margin = new System.Windows.Forms.Padding(2);
			this.Name = "frmAreaDepartHyp";
			this.Text = "QUẢN LÝ BỘ PHẬN";
			this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
			this.Load += new System.EventHandler(this.frmAreaDepartHyp_Load);
			this.toolStrip1.ResumeLayout(false);
			this.toolStrip1.PerformLayout();
			((System.ComponentModel.ISupportInitialize)(this.gridView1)).EndInit();
			this.splitContainer1.Panel1.ResumeLayout(false);
			this.splitContainer1.Panel1.PerformLayout();
			this.splitContainer1.Panel2.ResumeLayout(false);
			this.splitContainer1.Panel2.PerformLayout();
			((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
			this.splitContainer1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.dtgvArea)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gvArea)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtgvDepartment)).EndInit();
			this.contextMenuStrip1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.gvDepartment)).EndInit();
			this.panel6.ResumeLayout(false);
			this.panel6.PerformLayout();
			this.ResumeLayout(false);

		}

		#endregion

		private System.Windows.Forms.ToolStrip toolStrip1;
		private System.Windows.Forms.ToolStripButton btnCreateDep;
		private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
		private System.Windows.Forms.ToolStripButton btnEditDep;
		private System.Windows.Forms.ToolStripSeparator toolStripSeparator7;
		private System.Windows.Forms.ToolStripButton btnDelDep;
		private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
		private System.Windows.Forms.ToolStripButton btnRefresh;
		private System.Windows.Forms.ToolStripSeparator toolStripSeparator9;
		private System.Windows.Forms.ToolStripButton btnCreateArea;
		private System.Windows.Forms.ToolStripSeparator toolStripSeparator6;
		private System.Windows.Forms.ToolStripButton btnEditArea;
		private System.Windows.Forms.ToolStripSeparator toolStripSeparator8;
		private System.Windows.Forms.ToolStripButton btnDelArea;
		private DevExpress.XtraGrid.Views.Grid.GridView gridView1;
		private System.Windows.Forms.SplitContainer splitContainer1;
		private System.Windows.Forms.CheckBox cbShowAll;
		private System.Windows.Forms.Panel panel6;
		private DevExpress.XtraEditors.SimpleButton btnCancelSearch;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.TextBox txbSearch;
		private DevExpress.XtraEditors.SimpleButton btnSearch;
		private DevExpress.XtraEditors.LabelControl labelControl1;
		private DevExpress.XtraGrid.GridControl dtgvArea;
		private DevExpress.XtraGrid.Views.Grid.GridView gvArea;
		private DevExpress.XtraGrid.Columns.GridColumn colIDAr;
		private DevExpress.XtraGrid.Columns.GridColumn colAreaCodeAr;
		private DevExpress.XtraGrid.Columns.GridColumn colManagerIDAr;
		private DevExpress.XtraGrid.Columns.GridColumn colManagerCodeAr;
		private DevExpress.XtraGrid.Columns.GridColumn gridColumn2;
		private DevExpress.XtraGrid.GridControl dtgvDepartment;
		private DevExpress.XtraGrid.Views.Grid.GridView gvDepartment;
		private DevExpress.XtraGrid.Columns.GridColumn colIDDep;
		private DevExpress.XtraGrid.Columns.GridColumn colDepartmentCodeDep;
		private DevExpress.XtraGrid.Columns.GridColumn gridColumn3;
		private DevExpress.XtraGrid.Columns.GridColumn gridColumn4;
		private DevExpress.XtraGrid.Columns.GridColumn colSTT;
		private DevExpress.XtraGrid.Columns.GridColumn colAreaIDDep;
		private DevExpress.XtraGrid.Columns.GridColumn gridColumn5;
		private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
		private System.Windows.Forms.ToolStripMenuItem btnCheckStage;
		private DevExpress.XtraGrid.Columns.GridColumn colAreaName;
		private DevExpress.XtraGrid.Columns.GridColumn colDepartmentName;
		private DevExpress.XtraGrid.Columns.GridColumn gridColumn1;
	}
}