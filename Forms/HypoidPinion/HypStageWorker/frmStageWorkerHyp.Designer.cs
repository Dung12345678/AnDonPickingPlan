
namespace BMS
{
	partial class frmStageWorkerHyp
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
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmStageWorkerHyp));
			this.toolStrip1 = new System.Windows.Forms.ToolStrip();
			this.btnCreate = new System.Windows.Forms.ToolStripButton();
			this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
			this.btnEdit = new System.Windows.Forms.ToolStripButton();
			this.toolStripSeparator7 = new System.Windows.Forms.ToolStripSeparator();
			this.btnDel = new System.Windows.Forms.ToolStripButton();
			this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
			this.btnRefresh = new System.Windows.Forms.ToolStripButton();
			this.dtgvStageWorker = new DevExpress.XtraGrid.GridControl();
			this.gvStageWorker = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.colID = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colWorkerCode = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colStageID = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colStageCode = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colSTT = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colJobDescription = new DevExpress.XtraGrid.Columns.GridColumn();
			this.toolStrip1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.dtgvStageWorker)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gvStageWorker)).BeginInit();
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
            this.btnCreate,
            this.toolStripSeparator1,
            this.btnEdit,
            this.toolStripSeparator7,
            this.btnDel,
            this.toolStripSeparator2,
            this.btnRefresh});
			this.toolStrip1.LayoutStyle = System.Windows.Forms.ToolStripLayoutStyle.Flow;
			this.toolStrip1.Location = new System.Drawing.Point(0, 0);
			this.toolStrip1.Name = "toolStrip1";
			this.toolStrip1.Padding = new System.Windows.Forms.Padding(0, 0, 2, 0);
			this.toolStrip1.RenderMode = System.Windows.Forms.ToolStripRenderMode.Professional;
			this.toolStrip1.Size = new System.Drawing.Size(970, 42);
			this.toolStrip1.TabIndex = 31;
			this.toolStrip1.Text = "toolStrip2";
			// 
			// btnCreate
			// 
			this.btnCreate.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnCreate.ForeColor = System.Drawing.Color.Black;
			this.btnCreate.Image = ((System.Drawing.Image)(resources.GetObject("btnCreate.Image")));
			this.btnCreate.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnCreate.Name = "btnCreate";
			this.btnCreate.Size = new System.Drawing.Size(133, 41);
			this.btnCreate.Tag = "frmProduct_AddProductH";
			this.btnCreate.Text = "Thêm công nhân";
			this.btnCreate.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnCreate.Click += new System.EventHandler(this.btnCreate_Click);
			// 
			// toolStripSeparator1
			// 
			this.toolStripSeparator1.Name = "toolStripSeparator1";
			this.toolStripSeparator1.Size = new System.Drawing.Size(6, 23);
			// 
			// btnEdit
			// 
			this.btnEdit.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnEdit.ForeColor = System.Drawing.Color.Black;
			this.btnEdit.Image = ((System.Drawing.Image)(resources.GetObject("btnEdit.Image")));
			this.btnEdit.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnEdit.Name = "btnEdit";
			this.btnEdit.Size = new System.Drawing.Size(121, 41);
			this.btnEdit.Tag = "frmProduct_EditProductH";
			this.btnEdit.Text = "Sửa công nhân";
			this.btnEdit.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnEdit.Click += new System.EventHandler(this.btnEdit_Click);
			// 
			// toolStripSeparator7
			// 
			this.toolStripSeparator7.Name = "toolStripSeparator7";
			this.toolStripSeparator7.Size = new System.Drawing.Size(6, 23);
			// 
			// btnDel
			// 
			this.btnDel.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnDel.ForeColor = System.Drawing.Color.Black;
			this.btnDel.Image = ((System.Drawing.Image)(resources.GetObject("btnDel.Image")));
			this.btnDel.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnDel.Name = "btnDel";
			this.btnDel.Size = new System.Drawing.Size(121, 41);
			this.btnDel.Tag = "frmProduct_DeleteProductH";
			this.btnDel.Text = "Xóa công nhân";
			this.btnDel.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnDel.Click += new System.EventHandler(this.btnDel_Click);
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
			// dtgvStageWorker
			// 
			this.dtgvStageWorker.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.dtgvStageWorker.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.dtgvStageWorker.Location = new System.Drawing.Point(3, 46);
			this.dtgvStageWorker.LookAndFeel.Style = DevExpress.LookAndFeel.LookAndFeelStyle.Flat;
			this.dtgvStageWorker.MainView = this.gvStageWorker;
			this.dtgvStageWorker.Name = "dtgvStageWorker";
			this.dtgvStageWorker.Size = new System.Drawing.Size(964, 484);
			this.dtgvStageWorker.TabIndex = 38;
			this.dtgvStageWorker.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.gvStageWorker});
			// 
			// gvStageWorker
			// 
			this.gvStageWorker.ColumnPanelRowHeight = 40;
			this.gvStageWorker.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
            this.colID,
            this.colWorkerCode,
            this.colStageID,
            this.colStageCode,
            this.colSTT,
            this.colJobDescription});
			this.gvStageWorker.GridControl = this.dtgvStageWorker;
			this.gvStageWorker.Name = "gvStageWorker";
			this.gvStageWorker.OptionsBehavior.AllowAddRows = DevExpress.Utils.DefaultBoolean.False;
			this.gvStageWorker.OptionsBehavior.AllowDeleteRows = DevExpress.Utils.DefaultBoolean.False;
			this.gvStageWorker.OptionsBehavior.AllowFixedGroups = DevExpress.Utils.DefaultBoolean.False;
			this.gvStageWorker.OptionsBehavior.AllowIncrementalSearch = true;
			this.gvStageWorker.OptionsBehavior.Editable = false;
			this.gvStageWorker.OptionsCustomization.AllowColumnMoving = false;
			this.gvStageWorker.OptionsCustomization.AllowQuickHideColumns = false;
			this.gvStageWorker.OptionsFind.AlwaysVisible = true;
			this.gvStageWorker.OptionsFind.ShowCloseButton = false;
			this.gvStageWorker.OptionsView.ShowGroupPanel = false;
			this.gvStageWorker.RowHeight = 25;
			this.gvStageWorker.ScrollStyle = DevExpress.XtraGrid.Views.Grid.ScrollStyleFlags.LiveVertScroll;
			this.gvStageWorker.DoubleClick += new System.EventHandler(this.gvStageWorker_DoubleClick);
			// 
			// colID
			// 
			this.colID.Caption = "ID";
			this.colID.FieldName = "ID";
			this.colID.Name = "colID";
			this.colID.OptionsColumn.AllowEdit = false;
			// 
			// colWorkerCode
			// 
			this.colWorkerCode.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.colWorkerCode.AppearanceCell.Options.UseFont = true;
			this.colWorkerCode.AppearanceCell.Options.UseTextOptions = true;
			this.colWorkerCode.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colWorkerCode.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colWorkerCode.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colWorkerCode.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colWorkerCode.AppearanceHeader.ForeColor = System.Drawing.Color.Black;
			this.colWorkerCode.AppearanceHeader.Options.UseFont = true;
			this.colWorkerCode.AppearanceHeader.Options.UseForeColor = true;
			this.colWorkerCode.AppearanceHeader.Options.UseTextOptions = true;
			this.colWorkerCode.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colWorkerCode.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colWorkerCode.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colWorkerCode.Caption = "MÃ NHÂN VIÊN";
			this.colWorkerCode.FieldName = "WorkerCode";
			this.colWorkerCode.Name = "colWorkerCode";
			this.colWorkerCode.OptionsColumn.AllowEdit = false;
			this.colWorkerCode.Visible = true;
			this.colWorkerCode.VisibleIndex = 1;
			this.colWorkerCode.Width = 330;
			// 
			// colStageID
			// 
			this.colStageID.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.colStageID.AppearanceCell.Options.UseFont = true;
			this.colStageID.AppearanceCell.Options.UseTextOptions = true;
			this.colStageID.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colStageID.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colStageID.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colStageID.AppearanceHeader.ForeColor = System.Drawing.Color.Black;
			this.colStageID.AppearanceHeader.Options.UseFont = true;
			this.colStageID.AppearanceHeader.Options.UseForeColor = true;
			this.colStageID.AppearanceHeader.Options.UseTextOptions = true;
			this.colStageID.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colStageID.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colStageID.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colStageID.Caption = "Mô tả";
			this.colStageID.FieldName = "StageID";
			this.colStageID.Name = "colStageID";
			this.colStageID.OptionsColumn.AllowEdit = false;
			this.colStageID.Width = 338;
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
			this.colStageCode.VisibleIndex = 2;
			this.colStageCode.Width = 339;
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
			this.colSTT.Visible = true;
			this.colSTT.VisibleIndex = 0;
			this.colSTT.Width = 80;
			// 
			// colJobDescription
			// 
			this.colJobDescription.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.colJobDescription.AppearanceCell.Options.UseFont = true;
			this.colJobDescription.AppearanceCell.Options.UseTextOptions = true;
			this.colJobDescription.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colJobDescription.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colJobDescription.AppearanceHeader.Options.UseFont = true;
			this.colJobDescription.AppearanceHeader.Options.UseTextOptions = true;
			this.colJobDescription.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colJobDescription.Caption = "MÔ TẢ CÔNG VIỆC";
			this.colJobDescription.FieldName = "JobDescription";
			this.colJobDescription.Name = "colJobDescription";
			this.colJobDescription.Visible = true;
			this.colJobDescription.VisibleIndex = 3;
			this.colJobDescription.Width = 197;
			// 
			// frmStageWorkerHyp
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(970, 533);
			this.Controls.Add(this.dtgvStageWorker);
			this.Controls.Add(this.toolStrip1);
			this.Margin = new System.Windows.Forms.Padding(2);
			this.Name = "frmStageWorkerHyp";
			this.Text = "QUẢN LÝ NHÂN VIÊN CÔNG ĐOẠN";
			this.Load += new System.EventHandler(this.frmStageWorkerHyp_Load);
			this.toolStrip1.ResumeLayout(false);
			this.toolStrip1.PerformLayout();
			((System.ComponentModel.ISupportInitialize)(this.dtgvStageWorker)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gvStageWorker)).EndInit();
			this.ResumeLayout(false);

		}

		#endregion

		private System.Windows.Forms.ToolStrip toolStrip1;
		private System.Windows.Forms.ToolStripButton btnCreate;
		private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
		private System.Windows.Forms.ToolStripButton btnEdit;
		private System.Windows.Forms.ToolStripSeparator toolStripSeparator7;
		private System.Windows.Forms.ToolStripButton btnDel;
		private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
		private System.Windows.Forms.ToolStripButton btnRefresh;
		private DevExpress.XtraGrid.GridControl dtgvStageWorker;
		private DevExpress.XtraGrid.Views.Grid.GridView gvStageWorker;
		private DevExpress.XtraGrid.Columns.GridColumn colID;
		private DevExpress.XtraGrid.Columns.GridColumn colWorkerCode;
		private DevExpress.XtraGrid.Columns.GridColumn colStageID;
		private DevExpress.XtraGrid.Columns.GridColumn colStageCode;
		private DevExpress.XtraGrid.Columns.GridColumn colSTT;
		private DevExpress.XtraGrid.Columns.GridColumn colJobDescription;
	}
}