
namespace BMS
{
    partial class frmPart
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
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmPart));
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
			this.colPartCode = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colPartName = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colCreateDate = new DevExpress.XtraGrid.Columns.GridColumn();
			this.timer1 = new System.Windows.Forms.Timer(this.components);
			this.panel2 = new System.Windows.Forms.Panel();
			this.btnSearch = new DevExpress.XtraEditors.SimpleButton();
			this.txbSearch = new System.Windows.Forms.TextBox();
			this.dtpFrom = new System.Windows.Forms.DateTimePicker();
			this.dtpTo = new System.Windows.Forms.DateTimePicker();
			this.label1 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
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
			this.toolStrip1.Size = new System.Drawing.Size(953, 48);
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
			this.btnCreatePlan.Size = new System.Drawing.Size(119, 42);
			this.btnCreatePlan.Tag = "frmProduct_AddProductH";
			this.btnCreatePlan.Text = "Thêm bộ phận";
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
			this.btnEditPlan.Size = new System.Drawing.Size(106, 42);
			this.btnEditPlan.Tag = "frmProduct_EditProductH";
			this.btnEditPlan.Text = "Sửa bộ phận";
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
			this.btnDelPlan.Size = new System.Drawing.Size(107, 42);
			this.btnDelPlan.Tag = "frmProduct_DeleteProductH";
			this.btnDelPlan.Text = "Xóa bộ phận";
			this.btnDelPlan.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnDelPlan.Click += new System.EventHandler(this.btnDelPlan_Click);
			// 
			// toolStripSeparator8
			// 
			this.toolStripSeparator8.AutoSize = false;
			this.toolStripSeparator8.Name = "toolStripSeparator8";
			this.toolStripSeparator8.Size = new System.Drawing.Size(6, 40);
			this.toolStripSeparator8.Visible = false;
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
			this.btnExportExcel.Visible = false;
			this.btnExportExcel.Click += new System.EventHandler(this.btnExportExcel_Click);
			// 
			// toolStripSeparator11
			// 
			this.toolStripSeparator11.Name = "toolStripSeparator11";
			this.toolStripSeparator11.Size = new System.Drawing.Size(6, 23);
			this.toolStripSeparator11.Visible = false;
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
			this.btnImportExcel.Visible = false;
			this.btnImportExcel.Click += new System.EventHandler(this.btnImportExcel_Click);
			// 
			// toolStripSeparator12
			// 
			this.toolStripSeparator12.AutoSize = false;
			this.toolStripSeparator12.Name = "toolStripSeparator12";
			this.toolStripSeparator12.Size = new System.Drawing.Size(6, 40);
			this.toolStripSeparator12.Visible = false;
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
			this.grdData.Size = new System.Drawing.Size(953, 370);
			this.grdData.TabIndex = 29;
			this.grdData.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.grvData});
			// 
			// grvData
			// 
			this.grvData.ColumnPanelRowHeight = 60;
			this.grvData.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
            this.colID,
            this.colPartCode,
            this.colPartName,
            this.colCreateDate});
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
			this.colID.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colID.AppearanceHeader.ForeColor = System.Drawing.Color.Black;
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
			// colPartCode
			// 
			this.colPartCode.AppearanceCell.BackColor = System.Drawing.Color.White;
			this.colPartCode.AppearanceCell.BackColor2 = System.Drawing.Color.White;
			this.colPartCode.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.colPartCode.AppearanceCell.Options.UseBackColor = true;
			this.colPartCode.AppearanceCell.Options.UseFont = true;
			this.colPartCode.AppearanceCell.Options.UseTextOptions = true;
			this.colPartCode.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colPartCode.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colPartCode.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colPartCode.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colPartCode.AppearanceHeader.ForeColor = System.Drawing.Color.Black;
			this.colPartCode.AppearanceHeader.Options.UseFont = true;
			this.colPartCode.AppearanceHeader.Options.UseForeColor = true;
			this.colPartCode.AppearanceHeader.Options.UseTextOptions = true;
			this.colPartCode.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colPartCode.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colPartCode.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colPartCode.Caption = "Mã bộ phận";
			this.colPartCode.FieldName = "PartCode";
			this.colPartCode.MinWidth = 25;
			this.colPartCode.Name = "colPartCode";
			this.colPartCode.OptionsColumn.AllowEdit = false;
			this.colPartCode.Visible = true;
			this.colPartCode.VisibleIndex = 0;
			this.colPartCode.Width = 200;
			// 
			// colPartName
			// 
			this.colPartName.AppearanceCell.BackColor = System.Drawing.Color.White;
			this.colPartName.AppearanceCell.BackColor2 = System.Drawing.Color.White;
			this.colPartName.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.colPartName.AppearanceCell.Options.UseBackColor = true;
			this.colPartName.AppearanceCell.Options.UseFont = true;
			this.colPartName.AppearanceCell.Options.UseTextOptions = true;
			this.colPartName.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colPartName.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colPartName.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colPartName.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colPartName.AppearanceHeader.ForeColor = System.Drawing.Color.Black;
			this.colPartName.AppearanceHeader.Options.UseFont = true;
			this.colPartName.AppearanceHeader.Options.UseForeColor = true;
			this.colPartName.AppearanceHeader.Options.UseTextOptions = true;
			this.colPartName.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colPartName.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colPartName.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colPartName.Caption = "Tên bộ phận";
			this.colPartName.FieldName = "PartName";
			this.colPartName.MinWidth = 25;
			this.colPartName.Name = "colPartName";
			this.colPartName.OptionsColumn.AllowEdit = false;
			this.colPartName.Visible = true;
			this.colPartName.VisibleIndex = 1;
			this.colPartName.Width = 236;
			// 
			// colCreateDate
			// 
			this.colCreateDate.AppearanceCell.BackColor = System.Drawing.Color.White;
			this.colCreateDate.AppearanceCell.BackColor2 = System.Drawing.Color.White;
			this.colCreateDate.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.colCreateDate.AppearanceCell.Options.UseBackColor = true;
			this.colCreateDate.AppearanceCell.Options.UseFont = true;
			this.colCreateDate.AppearanceCell.Options.UseTextOptions = true;
			this.colCreateDate.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colCreateDate.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colCreateDate.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colCreateDate.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colCreateDate.AppearanceHeader.ForeColor = System.Drawing.Color.Black;
			this.colCreateDate.AppearanceHeader.Options.UseFont = true;
			this.colCreateDate.AppearanceHeader.Options.UseForeColor = true;
			this.colCreateDate.AppearanceHeader.Options.UseTextOptions = true;
			this.colCreateDate.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colCreateDate.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colCreateDate.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colCreateDate.Caption = "Ngày tạo";
			this.colCreateDate.FieldName = "CreateDate";
			this.colCreateDate.GroupFormat.FormatString = "dd-MM-yyyy";
			this.colCreateDate.GroupFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
			this.colCreateDate.MinWidth = 25;
			this.colCreateDate.Name = "colCreateDate";
			this.colCreateDate.OptionsColumn.AllowEdit = false;
			this.colCreateDate.Visible = true;
			this.colCreateDate.VisibleIndex = 2;
			this.colCreateDate.Width = 179;
			// 
			// panel2
			// 
			this.panel2.Controls.Add(this.btnSearch);
			this.panel2.Controls.Add(this.txbSearch);
			this.panel2.Controls.Add(this.dtpFrom);
			this.panel2.Controls.Add(this.dtpTo);
			this.panel2.Controls.Add(this.label1);
			this.panel2.Controls.Add(this.label3);
			this.panel2.Location = new System.Drawing.Point(0, 43);
			this.panel2.Name = "panel2";
			this.panel2.Size = new System.Drawing.Size(627, 34);
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
			// frmPart
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(953, 450);
			this.Controls.Add(this.panel2);
			this.Controls.Add(this.grdData);
			this.Controls.Add(this.toolStrip1);
			this.Margin = new System.Windows.Forms.Padding(4);
			this.Name = "frmPart";
			this.Text = "DANH SÁCH BỘ PHẬN";
			this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
			this.Load += new System.EventHandler(this.frmAssemblyOrder_Load);
			this.toolStrip1.ResumeLayout(false);
			this.toolStrip1.PerformLayout();
			((System.ComponentModel.ISupportInitialize)(this.grdData)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grvData)).EndInit();
			this.panel2.ResumeLayout(false);
			this.panel2.PerformLayout();
			this.ResumeLayout(false);

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
        private DevExpress.XtraGrid.Columns.GridColumn colPartCode;
        private DevExpress.XtraGrid.Columns.GridColumn colPartName;
        private DevExpress.XtraGrid.Columns.GridColumn colCreateDate;
		private System.Windows.Forms.Timer timer1;
		private System.Windows.Forms.Panel panel2;
		private DevExpress.XtraEditors.SimpleButton btnSearch;
		private System.Windows.Forms.TextBox txbSearch;
		private System.Windows.Forms.DateTimePicker dtpFrom;
		private System.Windows.Forms.DateTimePicker dtpTo;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label3;
	}
}