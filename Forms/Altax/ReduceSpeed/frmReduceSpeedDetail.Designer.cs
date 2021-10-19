namespace BMS
{
	partial class frmReduceSpeedDetail
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
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmReduceSpeedDetail));
			this.mnuMenu = new System.Windows.Forms.ToolStrip();
			this.btnSave = new System.Windows.Forms.ToolStripButton();
			this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
			this.btnSaveNew = new System.Windows.Forms.ToolStripButton();
			this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
			this.saveToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.saveNewToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.closeToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.NMin = new System.Windows.Forms.NumericUpDown();
			this.txtStandValue = new System.Windows.Forms.TextBox();
			this.label14 = new System.Windows.Forms.Label();
			this.label10 = new System.Windows.Forms.Label();
			this.label8 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.txtName = new System.Windows.Forms.TextBox();
			this.label2 = new System.Windows.Forms.Label();
			this.NMax = new System.Windows.Forms.NumericUpDown();
			this.NTimeGetDataVongQuay = new System.Windows.Forms.NumericUpDown();
			this.NTimeRunF = new System.Windows.Forms.NumericUpDown();
			this.NudTimeTeaching = new System.Windows.Forms.NumericUpDown();
			this.NLine = new System.Windows.Forms.NumericUpDown();
			this.mnuMenu.SuspendLayout();
			this.contextMenuStrip1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.NMin)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.NMax)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.NTimeGetDataVongQuay)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.NTimeRunF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.NudTimeTeaching)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.NLine)).BeginInit();
			this.SuspendLayout();
			// 
			// mnuMenu
			// 
			this.mnuMenu.AutoSize = false;
			this.mnuMenu.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(234)))), ((int)(((byte)(234)))), ((int)(((byte)(234)))));
			this.mnuMenu.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.mnuMenu.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden;
			this.mnuMenu.ImageScalingSize = new System.Drawing.Size(20, 20);
			this.mnuMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.btnSave,
            this.toolStripSeparator2,
            this.btnSaveNew});
			this.mnuMenu.LayoutStyle = System.Windows.Forms.ToolStripLayoutStyle.Flow;
			this.mnuMenu.Location = new System.Drawing.Point(0, 0);
			this.mnuMenu.Name = "mnuMenu";
			this.mnuMenu.RenderMode = System.Windows.Forms.ToolStripRenderMode.Professional;
			this.mnuMenu.Size = new System.Drawing.Size(695, 44);
			this.mnuMenu.TabIndex = 1;
			this.mnuMenu.Text = "toolStrip2";
			// 
			// btnSave
			// 
			this.btnSave.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnSave.Image = ((System.Drawing.Image)(resources.GetObject("btnSave.Image")));
			this.btnSave.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnSave.Name = "btnSave";
			this.btnSave.Size = new System.Drawing.Size(75, 37);
			this.btnSave.Tag = "frmProductPlan_AddPlan";
			this.btnSave.Text = "Cất && Đóng";
			this.btnSave.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
			// 
			// toolStripSeparator2
			// 
			this.toolStripSeparator2.AutoSize = false;
			this.toolStripSeparator2.Name = "toolStripSeparator2";
			this.toolStripSeparator2.Size = new System.Drawing.Size(6, 40);
			// 
			// btnSaveNew
			// 
			this.btnSaveNew.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnSaveNew.Image = ((System.Drawing.Image)(resources.GetObject("btnSaveNew.Image")));
			this.btnSaveNew.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnSaveNew.Name = "btnSaveNew";
			this.btnSaveNew.Size = new System.Drawing.Size(101, 37);
			this.btnSaveNew.Tag = "frmProductPlan_AddPlan";
			this.btnSaveNew.Text = "Cất && Thêm mới";
			this.btnSaveNew.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnSaveNew.Click += new System.EventHandler(this.btnSaveNew_Click);
			// 
			// contextMenuStrip1
			// 
			this.contextMenuStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
			this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.saveToolStripMenuItem,
            this.saveNewToolStripMenuItem,
            this.closeToolStripMenuItem});
			this.contextMenuStrip1.Name = "contextMenuStrip1";
			this.contextMenuStrip1.Size = new System.Drawing.Size(165, 70);
			// 
			// saveToolStripMenuItem
			// 
			this.saveToolStripMenuItem.Name = "saveToolStripMenuItem";
			this.saveToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.S)));
			this.saveToolStripMenuItem.Size = new System.Drawing.Size(164, 22);
			this.saveToolStripMenuItem.Text = "save";
			this.saveToolStripMenuItem.Visible = false;
			this.saveToolStripMenuItem.Click += new System.EventHandler(this.saveToolStripMenuItem_Click);
			// 
			// saveNewToolStripMenuItem
			// 
			this.saveNewToolStripMenuItem.Name = "saveNewToolStripMenuItem";
			this.saveNewToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.N)));
			this.saveNewToolStripMenuItem.Size = new System.Drawing.Size(164, 22);
			this.saveNewToolStripMenuItem.Text = "saveNew";
			this.saveNewToolStripMenuItem.Visible = false;
			this.saveNewToolStripMenuItem.Click += new System.EventHandler(this.saveNewToolStripMenuItem_Click);
			// 
			// closeToolStripMenuItem
			// 
			this.closeToolStripMenuItem.Name = "closeToolStripMenuItem";
			this.closeToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.X)));
			this.closeToolStripMenuItem.Size = new System.Drawing.Size(164, 22);
			this.closeToolStripMenuItem.Text = "close";
			this.closeToolStripMenuItem.Visible = false;
			this.closeToolStripMenuItem.Click += new System.EventHandler(this.closeToolStripMenuItem_Click);
			// 
			// NMin
			// 
			this.NMin.DecimalPlaces = 2;
			this.NMin.Location = new System.Drawing.Point(550, 56);
			this.NMin.Maximum = new decimal(new int[] {
            1661992959,
            1808227885,
            5,
            0});
			this.NMin.Minimum = new decimal(new int[] {
            -1304428545,
            434162106,
            542,
            -2147483648});
			this.NMin.Name = "NMin";
			this.NMin.Size = new System.Drawing.Size(120, 22);
			this.NMin.TabIndex = 1;
			this.NMin.ValueChanged += new System.EventHandler(this.NMin_ValueChanged);
			// 
			// txtStandValue
			// 
			this.txtStandValue.Font = new System.Drawing.Font("Tahoma", 9F);
			this.txtStandValue.Location = new System.Drawing.Point(206, 93);
			this.txtStandValue.Name = "txtStandValue";
			this.txtStandValue.Size = new System.Drawing.Size(244, 22);
			this.txtStandValue.TabIndex = 4;
			// 
			// label14
			// 
			this.label14.AutoSize = true;
			this.label14.Font = new System.Drawing.Font("Tahoma", 9F);
			this.label14.Location = new System.Drawing.Point(12, 133);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(130, 14);
			this.label14.TabIndex = 117;
			this.label14.Text = "Thời gian chạy 1 chiều";
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.Font = new System.Drawing.Font("Tahoma", 9F);
			this.label10.Location = new System.Drawing.Point(494, 139);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(29, 14);
			this.label10.TabIndex = 115;
			this.label10.Text = "Line";
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Font = new System.Drawing.Font("Tahoma", 9F);
			this.label8.Location = new System.Drawing.Point(12, 180);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(147, 14);
			this.label8.TabIndex = 113;
			this.label8.Text = "Thời gian teaching sensor";
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Font = new System.Drawing.Font("Tahoma", 9F);
			this.label7.Location = new System.Drawing.Point(12, 224);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(177, 14);
			this.label7.TabIndex = 112;
			this.label7.Text = "Thời gian lấy dữ liệu vòng quay";
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Font = new System.Drawing.Font("Tahoma", 9F);
			this.label5.Location = new System.Drawing.Point(494, 97);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(28, 14);
			this.label5.TabIndex = 111;
			this.label5.Text = "Max";
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Font = new System.Drawing.Font("Tahoma", 9F);
			this.label4.Location = new System.Drawing.Point(12, 97);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(100, 14);
			this.label4.TabIndex = 110;
			this.label4.Text = "Giá trị tiêu chuẩn";
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Font = new System.Drawing.Font("Tahoma", 9F);
			this.label3.Location = new System.Drawing.Point(12, 60);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(80, 14);
			this.label3.TabIndex = 109;
			this.label3.Text = "Tên giảm tốc";
			// 
			// txtName
			// 
			this.txtName.Font = new System.Drawing.Font("Tahoma", 9F);
			this.txtName.Location = new System.Drawing.Point(206, 56);
			this.txtName.Name = "txtName";
			this.txtName.Size = new System.Drawing.Size(244, 22);
			this.txtName.TabIndex = 0;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Font = new System.Drawing.Font("Tahoma", 9F);
			this.label2.Location = new System.Drawing.Point(494, 60);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(25, 14);
			this.label2.TabIndex = 107;
			this.label2.Text = "Min";
			// 
			// NMax
			// 
			this.NMax.DecimalPlaces = 2;
			this.NMax.Location = new System.Drawing.Point(550, 93);
			this.NMax.Maximum = new decimal(new int[] {
            276447231,
            23283,
            0,
            0});
			this.NMax.Minimum = new decimal(new int[] {
            -1530494977,
            232830,
            0,
            -2147483648});
			this.NMax.Name = "NMax";
			this.NMax.Size = new System.Drawing.Size(120, 22);
			this.NMax.TabIndex = 2;
			this.NMax.ValueChanged += new System.EventHandler(this.NMax_ValueChanged);
			// 
			// NTimeGetDataVongQuay
			// 
			this.NTimeGetDataVongQuay.DecimalPlaces = 2;
			this.NTimeGetDataVongQuay.Location = new System.Drawing.Point(206, 216);
			this.NTimeGetDataVongQuay.Maximum = new decimal(new int[] {
            1874919423,
            2328306,
            0,
            0});
			this.NTimeGetDataVongQuay.Minimum = new decimal(new int[] {
            -1486618625,
            232830643,
            0,
            -2147483648});
			this.NTimeGetDataVongQuay.Name = "NTimeGetDataVongQuay";
			this.NTimeGetDataVongQuay.Size = new System.Drawing.Size(120, 22);
			this.NTimeGetDataVongQuay.TabIndex = 7;
			// 
			// NTimeRunF
			// 
			this.NTimeRunF.DecimalPlaces = 2;
			this.NTimeRunF.Location = new System.Drawing.Point(206, 133);
			this.NTimeRunF.Maximum = new decimal(new int[] {
            -1530494977,
            232830,
            0,
            0});
			this.NTimeRunF.Minimum = new decimal(new int[] {
            1874919423,
            2328306,
            0,
            -2147483648});
			this.NTimeRunF.Name = "NTimeRunF";
			this.NTimeRunF.Size = new System.Drawing.Size(120, 22);
			this.NTimeRunF.TabIndex = 5;
			// 
			// NudTimeTeaching
			// 
			this.NudTimeTeaching.DecimalPlaces = 2;
			this.NudTimeTeaching.Location = new System.Drawing.Point(206, 172);
			this.NudTimeTeaching.Maximum = new decimal(new int[] {
            1569325055,
            23283064,
            0,
            0});
			this.NudTimeTeaching.Minimum = new decimal(new int[] {
            -1981284353,
            -1966660860,
            0,
            -2147483648});
			this.NudTimeTeaching.Name = "NudTimeTeaching";
			this.NudTimeTeaching.Size = new System.Drawing.Size(120, 22);
			this.NudTimeTeaching.TabIndex = 6;
			// 
			// NLine
			// 
			this.NLine.DecimalPlaces = 2;
			this.NLine.Location = new System.Drawing.Point(550, 131);
			this.NLine.Name = "NLine";
			this.NLine.Size = new System.Drawing.Size(120, 22);
			this.NLine.TabIndex = 3;
			// 
			// frmReduceSpeedDetail
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 14F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(695, 256);
			this.Controls.Add(this.NLine);
			this.Controls.Add(this.NudTimeTeaching);
			this.Controls.Add(this.NTimeRunF);
			this.Controls.Add(this.NTimeGetDataVongQuay);
			this.Controls.Add(this.NMax);
			this.Controls.Add(this.NMin);
			this.Controls.Add(this.txtStandValue);
			this.Controls.Add(this.label14);
			this.Controls.Add(this.label10);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.txtName);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.mnuMenu);
			this.Font = new System.Drawing.Font("Tahoma", 9F);
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.Margin = new System.Windows.Forms.Padding(2);
			this.Name = "frmReduceSpeedDetail";
			this.Text = "CHI TIẾT ĐIỀU KIỆN GIẢM TỐC PHÒNG CÁCH ÂM";
			this.Load += new System.EventHandler(this.ProductionPlanDetail_Load);
			this.mnuMenu.ResumeLayout(false);
			this.mnuMenu.PerformLayout();
			this.contextMenuStrip1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.NMin)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.NMax)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.NTimeGetDataVongQuay)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.NTimeRunF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.NudTimeTeaching)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.NLine)).EndInit();
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion
		private System.Windows.Forms.ToolStrip mnuMenu;
		private System.Windows.Forms.ToolStripButton btnSave;
		private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
		private System.Windows.Forms.ToolStripButton btnSaveNew;
		private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
		private System.Windows.Forms.ToolStripMenuItem saveToolStripMenuItem;
		private System.Windows.Forms.ToolStripMenuItem saveNewToolStripMenuItem;
		private System.Windows.Forms.ToolStripMenuItem closeToolStripMenuItem;
		private System.Windows.Forms.NumericUpDown NMin;
		private System.Windows.Forms.TextBox txtStandValue;
		private System.Windows.Forms.Label label14;
		private System.Windows.Forms.Label label10;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.TextBox txtName;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.NumericUpDown NMax;
		private System.Windows.Forms.NumericUpDown NTimeGetDataVongQuay;
		private System.Windows.Forms.NumericUpDown NTimeRunF;
		private System.Windows.Forms.NumericUpDown NudTimeTeaching;
		private System.Windows.Forms.NumericUpDown NLine;
	}
}