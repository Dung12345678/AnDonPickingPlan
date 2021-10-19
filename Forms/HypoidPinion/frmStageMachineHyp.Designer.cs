
namespace BMS
{
	partial class frmStageMachineHyp
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
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmStageMachineHyp));
			this.toolStrip1 = new System.Windows.Forms.ToolStrip();
			this.btnCreateStage = new System.Windows.Forms.ToolStripButton();
			this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
			this.btnEditStage = new System.Windows.Forms.ToolStripButton();
			this.toolStripSeparator7 = new System.Windows.Forms.ToolStripSeparator();
			this.btnDelStage = new System.Windows.Forms.ToolStripButton();
			this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
			this.btnCreateMachine = new System.Windows.Forms.ToolStripButton();
			this.toolStripSeparator4 = new System.Windows.Forms.ToolStripSeparator();
			this.btnEditMachine = new System.Windows.Forms.ToolStripButton();
			this.toolStripSeparator5 = new System.Windows.Forms.ToolStripSeparator();
			this.btnDelMachine = new System.Windows.Forms.ToolStripButton();
			this.toolStripSeparator3 = new System.Windows.Forms.ToolStripSeparator();
			this.btnRefresh = new System.Windows.Forms.ToolStripButton();
			this.toolStripSeparator6 = new System.Windows.Forms.ToolStripSeparator();
			this.btnShowWorker = new System.Windows.Forms.ToolStripButton();
			this.gridView1 = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.gridView2 = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.splitContainer1 = new System.Windows.Forms.SplitContainer();
			this.dtgvStage = new DevExpress.XtraGrid.GridControl();
			this.gvStage = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.colIDStage = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colStageCode = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colManagerID = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colManagerCode = new DevExpress.XtraGrid.Columns.GridColumn();
			this.gridColumn1 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colDepartmentID = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colDepartmentCode = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colStageName = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colDescription = new DevExpress.XtraGrid.Columns.GridColumn();
			this.labelControl1 = new DevExpress.XtraEditors.LabelControl();
			this.dtgvMachine = new DevExpress.XtraGrid.GridControl();
			this.gvMachine = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.colIDMac = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colMachineCode = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colSTT = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colStageID = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colStageCodeMac = new DevExpress.XtraGrid.Columns.GridColumn();
			this.gridColumn2 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.gridColumn3 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.cbShowAll = new System.Windows.Forms.CheckBox();
			this.panel6 = new System.Windows.Forms.Panel();
			this.btnCancelSearch = new DevExpress.XtraEditors.SimpleButton();
			this.label1 = new System.Windows.Forms.Label();
			this.txbSearch = new System.Windows.Forms.TextBox();
			this.btnSearch = new DevExpress.XtraEditors.SimpleButton();
			this.toolStrip1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.gridView1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
			this.splitContainer1.Panel1.SuspendLayout();
			this.splitContainer1.Panel2.SuspendLayout();
			this.splitContainer1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.dtgvStage)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gvStage)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtgvMachine)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gvMachine)).BeginInit();
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
            this.btnCreateStage,
            this.toolStripSeparator1,
            this.btnEditStage,
            this.toolStripSeparator7,
            this.btnDelStage,
            this.toolStripSeparator2,
            this.btnCreateMachine,
            this.toolStripSeparator4,
            this.btnEditMachine,
            this.toolStripSeparator5,
            this.btnDelMachine,
            this.toolStripSeparator3,
            this.btnRefresh,
            this.toolStripSeparator6,
            this.btnShowWorker});
			this.toolStrip1.LayoutStyle = System.Windows.Forms.ToolStripLayoutStyle.Flow;
			this.toolStrip1.Location = new System.Drawing.Point(0, 0);
			this.toolStrip1.Name = "toolStrip1";
			this.toolStrip1.Padding = new System.Windows.Forms.Padding(0, 0, 2, 0);
			this.toolStrip1.RenderMode = System.Windows.Forms.ToolStripRenderMode.Professional;
			this.toolStrip1.Size = new System.Drawing.Size(1028, 42);
			this.toolStrip1.TabIndex = 31;
			this.toolStrip1.Text = "toolStrip2";
			// 
			// btnCreateStage
			// 
			this.btnCreateStage.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnCreateStage.ForeColor = System.Drawing.Color.Black;
			this.btnCreateStage.Image = ((System.Drawing.Image)(resources.GetObject("btnCreateStage.Image")));
			this.btnCreateStage.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnCreateStage.Name = "btnCreateStage";
			this.btnCreateStage.Size = new System.Drawing.Size(121, 41);
			this.btnCreateStage.Tag = "frmProduct_AddProductH";
			this.btnCreateStage.Text = "Tạo công đoạn";
			this.btnCreateStage.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnCreateStage.Click += new System.EventHandler(this.btnCreateStage_Click);
			// 
			// toolStripSeparator1
			// 
			this.toolStripSeparator1.Name = "toolStripSeparator1";
			this.toolStripSeparator1.Size = new System.Drawing.Size(6, 23);
			// 
			// btnEditStage
			// 
			this.btnEditStage.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnEditStage.ForeColor = System.Drawing.Color.Black;
			this.btnEditStage.Image = ((System.Drawing.Image)(resources.GetObject("btnEditStage.Image")));
			this.btnEditStage.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnEditStage.Name = "btnEditStage";
			this.btnEditStage.Size = new System.Drawing.Size(121, 41);
			this.btnEditStage.Tag = "frmProduct_EditProductH";
			this.btnEditStage.Text = "Sửa công đoạn";
			this.btnEditStage.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnEditStage.Click += new System.EventHandler(this.btnEditStage_Click);
			// 
			// toolStripSeparator7
			// 
			this.toolStripSeparator7.Name = "toolStripSeparator7";
			this.toolStripSeparator7.Size = new System.Drawing.Size(6, 23);
			// 
			// btnDelStage
			// 
			this.btnDelStage.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnDelStage.ForeColor = System.Drawing.Color.Black;
			this.btnDelStage.Image = ((System.Drawing.Image)(resources.GetObject("btnDelStage.Image")));
			this.btnDelStage.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnDelStage.Name = "btnDelStage";
			this.btnDelStage.Size = new System.Drawing.Size(121, 41);
			this.btnDelStage.Tag = "frmProduct_DeleteProductH";
			this.btnDelStage.Text = "Xóa công đoạn";
			this.btnDelStage.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnDelStage.Click += new System.EventHandler(this.btnDelStage_Click);
			// 
			// toolStripSeparator2
			// 
			this.toolStripSeparator2.AutoSize = false;
			this.toolStripSeparator2.Name = "toolStripSeparator2";
			this.toolStripSeparator2.Size = new System.Drawing.Size(6, 40);
			// 
			// btnCreateMachine
			// 
			this.btnCreateMachine.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnCreateMachine.ForeColor = System.Drawing.Color.Black;
			this.btnCreateMachine.Image = ((System.Drawing.Image)(resources.GetObject("btnCreateMachine.Image")));
			this.btnCreateMachine.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnCreateMachine.Name = "btnCreateMachine";
			this.btnCreateMachine.Size = new System.Drawing.Size(74, 41);
			this.btnCreateMachine.Tag = "frmProduct_AddProductH";
			this.btnCreateMachine.Text = "Tạo máy";
			this.btnCreateMachine.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnCreateMachine.Click += new System.EventHandler(this.btnCreateMachine_Click);
			// 
			// toolStripSeparator4
			// 
			this.toolStripSeparator4.Name = "toolStripSeparator4";
			this.toolStripSeparator4.Size = new System.Drawing.Size(6, 23);
			// 
			// btnEditMachine
			// 
			this.btnEditMachine.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnEditMachine.ForeColor = System.Drawing.Color.Black;
			this.btnEditMachine.Image = ((System.Drawing.Image)(resources.GetObject("btnEditMachine.Image")));
			this.btnEditMachine.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnEditMachine.Name = "btnEditMachine";
			this.btnEditMachine.Size = new System.Drawing.Size(74, 41);
			this.btnEditMachine.Tag = "frmProduct_EditProductH";
			this.btnEditMachine.Text = "Sửa máy";
			this.btnEditMachine.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnEditMachine.Click += new System.EventHandler(this.btnEditMachine_Click);
			// 
			// toolStripSeparator5
			// 
			this.toolStripSeparator5.Name = "toolStripSeparator5";
			this.toolStripSeparator5.Size = new System.Drawing.Size(6, 23);
			// 
			// btnDelMachine
			// 
			this.btnDelMachine.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnDelMachine.ForeColor = System.Drawing.Color.Black;
			this.btnDelMachine.Image = ((System.Drawing.Image)(resources.GetObject("btnDelMachine.Image")));
			this.btnDelMachine.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnDelMachine.Name = "btnDelMachine";
			this.btnDelMachine.Size = new System.Drawing.Size(74, 41);
			this.btnDelMachine.Tag = "frmProduct_DeleteProductH";
			this.btnDelMachine.Text = "Xóa máy";
			this.btnDelMachine.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnDelMachine.Click += new System.EventHandler(this.btnDelMachine_Click);
			// 
			// toolStripSeparator3
			// 
			this.toolStripSeparator3.AutoSize = false;
			this.toolStripSeparator3.Name = "toolStripSeparator3";
			this.toolStripSeparator3.Size = new System.Drawing.Size(6, 40);
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
			// toolStripSeparator6
			// 
			this.toolStripSeparator6.AutoSize = false;
			this.toolStripSeparator6.Name = "toolStripSeparator6";
			this.toolStripSeparator6.Size = new System.Drawing.Size(6, 40);
			// 
			// btnShowWorker
			// 
			this.btnShowWorker.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnShowWorker.ForeColor = System.Drawing.Color.Black;
			this.btnShowWorker.Image = ((System.Drawing.Image)(resources.GetObject("btnShowWorker.Image")));
			this.btnShowWorker.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnShowWorker.Name = "btnShowWorker";
			this.btnShowWorker.Size = new System.Drawing.Size(225, 41);
			this.btnShowWorker.Tag = "frmProduct_AddProductH";
			this.btnShowWorker.Text = "Quản lý nhân viên công đoạn";
			this.btnShowWorker.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnShowWorker.Click += new System.EventHandler(this.btnShowWorker_Click);
			// 
			// gridView1
			// 
			this.gridView1.Name = "gridView1";
			// 
			// gridView2
			// 
			this.gridView2.Name = "gridView2";
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
			this.splitContainer1.Panel1.Controls.Add(this.dtgvStage);
			this.splitContainer1.Panel1.Controls.Add(this.labelControl1);
			this.splitContainer1.Panel1MinSize = 15;
			// 
			// splitContainer1.Panel2
			// 
			this.splitContainer1.Panel2.Controls.Add(this.dtgvMachine);
			this.splitContainer1.Panel2.Controls.Add(this.cbShowAll);
			this.splitContainer1.Panel2.Controls.Add(this.panel6);
			this.splitContainer1.Size = new System.Drawing.Size(1052, 556);
			this.splitContainer1.SplitterDistance = 482;
			this.splitContainer1.SplitterWidth = 3;
			this.splitContainer1.TabIndex = 37;
			// 
			// dtgvStage
			// 
			this.dtgvStage.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.dtgvStage.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.dtgvStage.Location = new System.Drawing.Point(-1, 45);
			this.dtgvStage.LookAndFeel.Style = DevExpress.LookAndFeel.LookAndFeelStyle.Flat;
			this.dtgvStage.MainView = this.gvStage;
			this.dtgvStage.Name = "dtgvStage";
			this.dtgvStage.Size = new System.Drawing.Size(452, 506);
			this.dtgvStage.TabIndex = 38;
			this.dtgvStage.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.gvStage});
			// 
			// gvStage
			// 
			this.gvStage.ColumnPanelRowHeight = 40;
			this.gvStage.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
            this.colIDStage,
            this.colStageCode,
            this.colManagerID,
            this.colManagerCode,
            this.gridColumn1,
            this.colDepartmentID,
            this.colDepartmentCode,
            this.colStageName,
            this.colDescription});
			this.gvStage.GridControl = this.dtgvStage;
			this.gvStage.Name = "gvStage";
			this.gvStage.OptionsBehavior.AllowAddRows = DevExpress.Utils.DefaultBoolean.False;
			this.gvStage.OptionsBehavior.AllowDeleteRows = DevExpress.Utils.DefaultBoolean.False;
			this.gvStage.OptionsBehavior.AllowFixedGroups = DevExpress.Utils.DefaultBoolean.False;
			this.gvStage.OptionsBehavior.AllowIncrementalSearch = true;
			this.gvStage.OptionsBehavior.Editable = false;
			this.gvStage.OptionsCustomization.AllowColumnMoving = false;
			this.gvStage.OptionsCustomization.AllowQuickHideColumns = false;
			this.gvStage.OptionsFind.AlwaysVisible = true;
			this.gvStage.OptionsFind.ShowCloseButton = false;
			this.gvStage.OptionsView.ShowGroupPanel = false;
			this.gvStage.RowHeight = 25;
			this.gvStage.ScrollStyle = DevExpress.XtraGrid.Views.Grid.ScrollStyleFlags.LiveVertScroll;
			this.gvStage.FocusedRowChanged += new DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventHandler(this.gvStage_FocusedRowChanged);
			this.gvStage.DoubleClick += new System.EventHandler(this.gvStage_DoubleClick);
			// 
			// colIDStage
			// 
			this.colIDStage.Caption = "ID";
			this.colIDStage.FieldName = "ID";
			this.colIDStage.Name = "colIDStage";
			this.colIDStage.OptionsColumn.AllowEdit = false;
			// 
			// colStageCode
			// 
			this.colStageCode.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.colStageCode.AppearanceCell.Options.UseFont = true;
			this.colStageCode.AppearanceCell.Options.UseTextOptions = true;
			this.colStageCode.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colStageCode.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colStageCode.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colStageCode.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colStageCode.AppearanceHeader.ForeColor = System.Drawing.Color.Black;
			this.colStageCode.AppearanceHeader.Options.UseFont = true;
			this.colStageCode.AppearanceHeader.Options.UseForeColor = true;
			this.colStageCode.AppearanceHeader.Options.UseTextOptions = true;
			this.colStageCode.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colStageCode.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colStageCode.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colStageCode.Caption = "MÃ CÔNG ĐOẠN";
			this.colStageCode.FieldName = "StageCode";
			this.colStageCode.Name = "colStageCode";
			this.colStageCode.OptionsColumn.AllowEdit = false;
			this.colStageCode.Visible = true;
			this.colStageCode.VisibleIndex = 0;
			this.colStageCode.Width = 134;
			// 
			// colManagerID
			// 
			this.colManagerID.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.colManagerID.AppearanceCell.Options.UseFont = true;
			this.colManagerID.AppearanceCell.Options.UseTextOptions = true;
			this.colManagerID.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colManagerID.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colManagerID.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colManagerID.AppearanceHeader.ForeColor = System.Drawing.Color.Black;
			this.colManagerID.AppearanceHeader.Options.UseFont = true;
			this.colManagerID.AppearanceHeader.Options.UseForeColor = true;
			this.colManagerID.AppearanceHeader.Options.UseTextOptions = true;
			this.colManagerID.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colManagerID.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colManagerID.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colManagerID.Caption = "Mô tả";
			this.colManagerID.FieldName = "ManagerID";
			this.colManagerID.Name = "colManagerID";
			this.colManagerID.OptionsColumn.AllowEdit = false;
			this.colManagerID.Width = 338;
			// 
			// colManagerCode
			// 
			this.colManagerCode.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.colManagerCode.AppearanceCell.Options.UseFont = true;
			this.colManagerCode.AppearanceCell.Options.UseTextOptions = true;
			this.colManagerCode.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colManagerCode.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colManagerCode.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colManagerCode.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colManagerCode.AppearanceHeader.ForeColor = System.Drawing.Color.Black;
			this.colManagerCode.AppearanceHeader.Options.UseFont = true;
			this.colManagerCode.AppearanceHeader.Options.UseForeColor = true;
			this.colManagerCode.AppearanceHeader.Options.UseTextOptions = true;
			this.colManagerCode.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colManagerCode.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colManagerCode.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colManagerCode.Caption = "NGƯỜI PHỤ TRÁCH";
			this.colManagerCode.FieldName = "ManagerCode";
			this.colManagerCode.Name = "colManagerCode";
			this.colManagerCode.OptionsColumn.AllowEdit = false;
			this.colManagerCode.Visible = true;
			this.colManagerCode.VisibleIndex = 2;
			this.colManagerCode.Width = 131;
			// 
			// gridColumn1
			// 
			this.gridColumn1.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.gridColumn1.AppearanceCell.Options.UseFont = true;
			this.gridColumn1.AppearanceCell.Options.UseTextOptions = true;
			this.gridColumn1.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.gridColumn1.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.gridColumn1.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.gridColumn1.AppearanceHeader.ForeColor = System.Drawing.Color.Black;
			this.gridColumn1.AppearanceHeader.Options.UseFont = true;
			this.gridColumn1.AppearanceHeader.Options.UseForeColor = true;
			this.gridColumn1.AppearanceHeader.Options.UseTextOptions = true;
			this.gridColumn1.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.gridColumn1.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.gridColumn1.Caption = "STT";
			this.gridColumn1.FieldName = "STT";
			this.gridColumn1.Name = "gridColumn1";
			this.gridColumn1.Width = 151;
			// 
			// colDepartmentID
			// 
			this.colDepartmentID.Caption = "gridColumn1";
			this.colDepartmentID.FieldName = "DepartmentID";
			this.colDepartmentID.Name = "colDepartmentID";
			// 
			// colDepartmentCode
			// 
			this.colDepartmentCode.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.colDepartmentCode.AppearanceCell.Options.UseFont = true;
			this.colDepartmentCode.AppearanceCell.Options.UseTextOptions = true;
			this.colDepartmentCode.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colDepartmentCode.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colDepartmentCode.AppearanceHeader.Options.UseFont = true;
			this.colDepartmentCode.AppearanceHeader.Options.UseTextOptions = true;
			this.colDepartmentCode.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colDepartmentCode.Caption = "MÃ BỘ PHẬN";
			this.colDepartmentCode.FieldName = "DepartmentCode";
			this.colDepartmentCode.Name = "colDepartmentCode";
			this.colDepartmentCode.Visible = true;
			this.colDepartmentCode.VisibleIndex = 1;
			this.colDepartmentCode.Width = 114;
			// 
			// colStageName
			// 
			this.colStageName.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.colStageName.AppearanceCell.Options.UseFont = true;
			this.colStageName.AppearanceCell.Options.UseTextOptions = true;
			this.colStageName.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colStageName.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colStageName.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colStageName.AppearanceHeader.Options.UseFont = true;
			this.colStageName.AppearanceHeader.Options.UseTextOptions = true;
			this.colStageName.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colStageName.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colStageName.Caption = "TÊN CÔNG ĐOẠN";
			this.colStageName.FieldName = "StageName";
			this.colStageName.Name = "colStageName";
			this.colStageName.Visible = true;
			this.colStageName.VisibleIndex = 3;
			this.colStageName.Width = 106;
			// 
			// colDescription
			// 
			this.colDescription.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.colDescription.AppearanceCell.Options.UseFont = true;
			this.colDescription.AppearanceCell.Options.UseTextOptions = true;
			this.colDescription.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colDescription.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colDescription.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colDescription.AppearanceHeader.Options.UseFont = true;
			this.colDescription.AppearanceHeader.Options.UseTextOptions = true;
			this.colDescription.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colDescription.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colDescription.Caption = "MÔ TẢ";
			this.colDescription.FieldName = "Description";
			this.colDescription.Name = "colDescription";
			this.colDescription.Visible = true;
			this.colDescription.VisibleIndex = 4;
			// 
			// labelControl1
			// 
			this.labelControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)));
			this.labelControl1.Appearance.Font = new System.Drawing.Font("Tahoma", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.labelControl1.Location = new System.Drawing.Point(156, 13);
			this.labelControl1.Name = "labelControl1";
			this.labelControl1.Size = new System.Drawing.Size(148, 17);
			this.labelControl1.TabIndex = 19;
			this.labelControl1.Text = "Danh sách công đoạn";
			// 
			// dtgvMachine
			// 
			this.dtgvMachine.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.dtgvMachine.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.dtgvMachine.Location = new System.Drawing.Point(3, 45);
			this.dtgvMachine.LookAndFeel.Style = DevExpress.LookAndFeel.LookAndFeelStyle.Flat;
			this.dtgvMachine.MainView = this.gvMachine;
			this.dtgvMachine.Name = "dtgvMachine";
			this.dtgvMachine.Size = new System.Drawing.Size(510, 506);
			this.dtgvMachine.TabIndex = 38;
			this.dtgvMachine.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.gvMachine});
			// 
			// gvMachine
			// 
			this.gvMachine.ColumnPanelRowHeight = 40;
			this.gvMachine.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
            this.colIDMac,
            this.colMachineCode,
            this.colSTT,
            this.colStageID,
            this.colStageCodeMac,
            this.gridColumn2,
            this.gridColumn3});
			this.gvMachine.GridControl = this.dtgvMachine;
			this.gvMachine.Name = "gvMachine";
			this.gvMachine.OptionsBehavior.AllowAddRows = DevExpress.Utils.DefaultBoolean.False;
			this.gvMachine.OptionsBehavior.AllowDeleteRows = DevExpress.Utils.DefaultBoolean.False;
			this.gvMachine.OptionsBehavior.AllowFixedGroups = DevExpress.Utils.DefaultBoolean.False;
			this.gvMachine.OptionsBehavior.AllowIncrementalSearch = true;
			this.gvMachine.OptionsBehavior.Editable = false;
			this.gvMachine.OptionsCustomization.AllowColumnMoving = false;
			this.gvMachine.OptionsCustomization.AllowQuickHideColumns = false;
			this.gvMachine.OptionsFind.AlwaysVisible = true;
			this.gvMachine.OptionsFind.ShowCloseButton = false;
			this.gvMachine.OptionsView.ShowGroupPanel = false;
			this.gvMachine.RowHeight = 25;
			this.gvMachine.ScrollStyle = DevExpress.XtraGrid.Views.Grid.ScrollStyleFlags.LiveVertScroll;
			this.gvMachine.DoubleClick += new System.EventHandler(this.gvMachine_DoubleClick);
			// 
			// colIDMac
			// 
			this.colIDMac.Caption = "ID";
			this.colIDMac.FieldName = "ID";
			this.colIDMac.Name = "colIDMac";
			this.colIDMac.OptionsColumn.AllowEdit = false;
			// 
			// colMachineCode
			// 
			this.colMachineCode.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.colMachineCode.AppearanceCell.Options.UseFont = true;
			this.colMachineCode.AppearanceCell.Options.UseTextOptions = true;
			this.colMachineCode.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colMachineCode.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colMachineCode.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colMachineCode.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colMachineCode.AppearanceHeader.ForeColor = System.Drawing.Color.Black;
			this.colMachineCode.AppearanceHeader.Options.UseFont = true;
			this.colMachineCode.AppearanceHeader.Options.UseForeColor = true;
			this.colMachineCode.AppearanceHeader.Options.UseTextOptions = true;
			this.colMachineCode.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colMachineCode.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colMachineCode.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colMachineCode.Caption = "MÃ MÁY";
			this.colMachineCode.FieldName = "MachineCode";
			this.colMachineCode.Name = "colMachineCode";
			this.colMachineCode.OptionsColumn.AllowEdit = false;
			this.colMachineCode.Visible = true;
			this.colMachineCode.VisibleIndex = 0;
			this.colMachineCode.Width = 130;
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
			this.colSTT.Width = 173;
			// 
			// colStageID
			// 
			this.colStageID.Caption = "gridColumn1";
			this.colStageID.FieldName = "StageID";
			this.colStageID.Name = "colStageID";
			// 
			// colStageCodeMac
			// 
			this.colStageCodeMac.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.colStageCodeMac.AppearanceCell.Options.UseFont = true;
			this.colStageCodeMac.AppearanceCell.Options.UseTextOptions = true;
			this.colStageCodeMac.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colStageCodeMac.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colStageCodeMac.AppearanceHeader.Options.UseFont = true;
			this.colStageCodeMac.AppearanceHeader.Options.UseTextOptions = true;
			this.colStageCodeMac.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colStageCodeMac.Caption = "MÃ CÔNG ĐOẠN";
			this.colStageCodeMac.FieldName = "StageCode";
			this.colStageCodeMac.Name = "colStageCodeMac";
			this.colStageCodeMac.Visible = true;
			this.colStageCodeMac.VisibleIndex = 2;
			this.colStageCodeMac.Width = 152;
			// 
			// gridColumn2
			// 
			this.gridColumn2.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.gridColumn2.AppearanceCell.Options.UseFont = true;
			this.gridColumn2.AppearanceCell.Options.UseTextOptions = true;
			this.gridColumn2.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.gridColumn2.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.gridColumn2.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.gridColumn2.AppearanceHeader.Options.UseFont = true;
			this.gridColumn2.AppearanceHeader.Options.UseTextOptions = true;
			this.gridColumn2.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.gridColumn2.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.gridColumn2.Caption = "TÊN MÁY";
			this.gridColumn2.FieldName = "MachineName";
			this.gridColumn2.Name = "gridColumn2";
			this.gridColumn2.Visible = true;
			this.gridColumn2.VisibleIndex = 1;
			this.gridColumn2.Width = 117;
			// 
			// gridColumn3
			// 
			this.gridColumn3.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.gridColumn3.AppearanceCell.Options.UseFont = true;
			this.gridColumn3.AppearanceCell.Options.UseTextOptions = true;
			this.gridColumn3.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.gridColumn3.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.gridColumn3.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.gridColumn3.AppearanceHeader.Options.UseFont = true;
			this.gridColumn3.AppearanceHeader.Options.UseTextOptions = true;
			this.gridColumn3.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.gridColumn3.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.gridColumn3.Caption = "MÔ TẢ";
			this.gridColumn3.FieldName = "Description";
			this.gridColumn3.Name = "gridColumn3";
			this.gridColumn3.Visible = true;
			this.gridColumn3.VisibleIndex = 3;
			this.gridColumn3.Width = 124;
			// 
			// cbShowAll
			// 
			this.cbShowAll.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.cbShowAll.AutoSize = true;
			this.cbShowAll.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.cbShowAll.Location = new System.Drawing.Point(414, 13);
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
			this.panel6.Size = new System.Drawing.Size(440, 36);
			this.panel6.TabIndex = 35;
			// 
			// btnCancelSearch
			// 
			this.btnCancelSearch.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelSearch.Location = new System.Drawing.Point(370, 6);
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
			this.txbSearch.Location = new System.Drawing.Point(69, 5);
			this.txbSearch.Name = "txbSearch";
			this.txbSearch.Size = new System.Drawing.Size(225, 24);
			this.txbSearch.TabIndex = 26;
			this.txbSearch.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txbSearch_KeyPress);
			// 
			// btnSearch
			// 
			this.btnSearch.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSearch.Location = new System.Drawing.Point(302, 6);
			this.btnSearch.Name = "btnSearch";
			this.btnSearch.Size = new System.Drawing.Size(63, 20);
			this.btnSearch.TabIndex = 25;
			this.btnSearch.Text = "Tìm kiếm";
			this.btnSearch.Click += new System.EventHandler(this.btnSearch_Click);
			// 
			// frmStageMachineHyp
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(1028, 602);
			this.Controls.Add(this.splitContainer1);
			this.Controls.Add(this.toolStrip1);
			this.Margin = new System.Windows.Forms.Padding(2);
			this.Name = "frmStageMachineHyp";
			this.Text = "QUẢN LÝ MÁY VÀ CÔNG ĐOẠN";
			this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
			this.Load += new System.EventHandler(this.frmStageMachineHyp_Load);
			this.toolStrip1.ResumeLayout(false);
			this.toolStrip1.PerformLayout();
			((System.ComponentModel.ISupportInitialize)(this.gridView1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView2)).EndInit();
			this.splitContainer1.Panel1.ResumeLayout(false);
			this.splitContainer1.Panel1.PerformLayout();
			this.splitContainer1.Panel2.ResumeLayout(false);
			this.splitContainer1.Panel2.PerformLayout();
			((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
			this.splitContainer1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.dtgvStage)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gvStage)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtgvMachine)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gvMachine)).EndInit();
			this.panel6.ResumeLayout(false);
			this.panel6.PerformLayout();
			this.ResumeLayout(false);

		}

		#endregion

		private System.Windows.Forms.ToolStrip toolStrip1;
		private System.Windows.Forms.ToolStripButton btnCreateStage;
		private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
		private System.Windows.Forms.ToolStripButton btnEditStage;
		private System.Windows.Forms.ToolStripSeparator toolStripSeparator7;
		private System.Windows.Forms.ToolStripButton btnDelStage;
		private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
		private System.Windows.Forms.ToolStripButton btnRefresh;
		private System.Windows.Forms.ToolStripSeparator toolStripSeparator3;
		private System.Windows.Forms.ToolStripButton btnCreateMachine;
		private System.Windows.Forms.ToolStripSeparator toolStripSeparator4;
		private System.Windows.Forms.ToolStripButton btnEditMachine;
		private System.Windows.Forms.ToolStripSeparator toolStripSeparator5;
		private System.Windows.Forms.ToolStripButton btnDelMachine;
		private DevExpress.XtraGrid.Views.Grid.GridView gridView1;
		private DevExpress.XtraGrid.Views.Grid.GridView gridView2;
		private System.Windows.Forms.SplitContainer splitContainer1;
		private System.Windows.Forms.CheckBox cbShowAll;
		private System.Windows.Forms.Panel panel6;
		private DevExpress.XtraEditors.SimpleButton btnCancelSearch;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.TextBox txbSearch;
		private DevExpress.XtraEditors.SimpleButton btnSearch;
		private DevExpress.XtraGrid.GridControl dtgvStage;
		private DevExpress.XtraGrid.Views.Grid.GridView gvStage;
		private DevExpress.XtraGrid.Columns.GridColumn colIDStage;
		private DevExpress.XtraGrid.Columns.GridColumn colStageCode;
		private DevExpress.XtraGrid.Columns.GridColumn colManagerID;
		private DevExpress.XtraGrid.Columns.GridColumn colManagerCode;
		private DevExpress.XtraGrid.Columns.GridColumn gridColumn1;
		private DevExpress.XtraGrid.Columns.GridColumn colDepartmentID;
		private DevExpress.XtraGrid.Columns.GridColumn colDepartmentCode;
		private DevExpress.XtraEditors.LabelControl labelControl1;
		private DevExpress.XtraGrid.GridControl dtgvMachine;
		private DevExpress.XtraGrid.Views.Grid.GridView gvMachine;
		private DevExpress.XtraGrid.Columns.GridColumn colIDMac;
		private DevExpress.XtraGrid.Columns.GridColumn colMachineCode;
		private DevExpress.XtraGrid.Columns.GridColumn colSTT;
		private DevExpress.XtraGrid.Columns.GridColumn colStageID;
		private DevExpress.XtraGrid.Columns.GridColumn colStageCodeMac;
		private System.Windows.Forms.ToolStripSeparator toolStripSeparator6;
		private System.Windows.Forms.ToolStripButton btnShowWorker;
		private DevExpress.XtraGrid.Columns.GridColumn colStageName;
		private DevExpress.XtraGrid.Columns.GridColumn colDescription;
		private DevExpress.XtraGrid.Columns.GridColumn gridColumn2;
		private DevExpress.XtraGrid.Columns.GridColumn gridColumn3;
	}
}