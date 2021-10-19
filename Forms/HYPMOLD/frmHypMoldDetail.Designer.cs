
namespace BMS
{
    partial class frmHypMoldDetail
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
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmHypMoldDetail));
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.cboType = new System.Windows.Forms.ComboBox();
			this.label7 = new System.Windows.Forms.Label();
			this.txtCheckingName = new System.Windows.Forms.TextBox();
			this.txtStt = new System.Windows.Forms.TextBox();
			this.txtStandardValue = new System.Windows.Forms.TextBox();
			this.txtMin = new System.Windows.Forms.TextBox();
			this.txtMax = new System.Windows.Forms.TextBox();
			this.txtUnit = new System.Windows.Forms.TextBox();
			this.mnuMenu = new System.Windows.Forms.ToolStrip();
			this.btnSaveClose = new System.Windows.Forms.ToolStripButton();
			this.txtMahang = new System.Windows.Forms.TextBox();
			this.label8 = new System.Windows.Forms.Label();
			this.checkBox1 = new System.Windows.Forms.CheckBox();
			this.txtCreateby = new System.Windows.Forms.TextBox();
			this.lbl = new System.Windows.Forms.Label();
			this.menuStrip1 = new System.Windows.Forms.MenuStrip();
			this.cấtToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.txtY = new System.Windows.Forms.TextBox();
			this.txtX = new System.Windows.Forms.TextBox();
			this.label9 = new System.Windows.Forms.Label();
			this.label10 = new System.Windows.Forms.Label();
			this.mnuMenu.SuspendLayout();
			this.menuStrip1.SuspendLayout();
			this.SuspendLayout();
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(11, 203);
			this.label1.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(42, 26);
			this.label1.TabIndex = 0;
			this.label1.Text = "STT";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(11, 154);
			this.label2.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(126, 26);
			this.label2.TabIndex = 0;
			this.label2.Text = "Mục kiểm tra";
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(11, 350);
			this.label3.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(46, 26);
			this.label3.TabIndex = 1;
			this.label3.Text = "Min";
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(11, 398);
			this.label4.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(49, 26);
			this.label4.TabIndex = 1;
			this.label4.Text = "Max";
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(11, 302);
			this.label5.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(159, 26);
			this.label5.TabIndex = 2;
			this.label5.Text = "Giá trị tiêu chuẩn";
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(11, 251);
			this.label6.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(52, 26);
			this.label6.TabIndex = 3;
			this.label6.Text = "Type";
			// 
			// cboType
			// 
			this.cboType.FormattingEnabled = true;
			this.cboType.Items.AddRange(new object[] {
            "",
            "Min ~ Max",
            "Tiêu chuẩn"});
			this.cboType.Location = new System.Drawing.Point(173, 248);
			this.cboType.Name = "cboType";
			this.cboType.Size = new System.Drawing.Size(217, 34);
			this.cboType.TabIndex = 3;
			this.cboType.SelectedIndexChanged += new System.EventHandler(this.cboType_SelectedIndexChanged);
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(11, 446);
			this.label7.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(67, 26);
			this.label7.TabIndex = 5;
			this.label7.Text = "Đơn vị";
			// 
			// txtCheckingName
			// 
			this.txtCheckingName.Location = new System.Drawing.Point(173, 151);
			this.txtCheckingName.Name = "txtCheckingName";
			this.txtCheckingName.Size = new System.Drawing.Size(343, 33);
			this.txtCheckingName.TabIndex = 1;
			this.txtCheckingName.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCheckingName_KeyDown);
			// 
			// txtStt
			// 
			this.txtStt.Location = new System.Drawing.Point(173, 200);
			this.txtStt.Name = "txtStt";
			this.txtStt.Size = new System.Drawing.Size(217, 33);
			this.txtStt.TabIndex = 2;
			this.txtStt.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtStt_KeyDown);
			// 
			// txtStandardValue
			// 
			this.txtStandardValue.Location = new System.Drawing.Point(173, 299);
			this.txtStandardValue.Name = "txtStandardValue";
			this.txtStandardValue.Size = new System.Drawing.Size(217, 33);
			this.txtStandardValue.TabIndex = 4;
			this.txtStandardValue.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtStandardValue_KeyDown);
			// 
			// txtMin
			// 
			this.txtMin.Location = new System.Drawing.Point(173, 347);
			this.txtMin.Name = "txtMin";
			this.txtMin.Size = new System.Drawing.Size(217, 33);
			this.txtMin.TabIndex = 5;
			this.txtMin.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtMin_KeyDown);
			this.txtMin.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtMin_KeyPress);
			// 
			// txtMax
			// 
			this.txtMax.Location = new System.Drawing.Point(173, 395);
			this.txtMax.Name = "txtMax";
			this.txtMax.Size = new System.Drawing.Size(217, 33);
			this.txtMax.TabIndex = 6;
			this.txtMax.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtMax_KeyDown);
			this.txtMax.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtMax_KeyPress);
			// 
			// txtUnit
			// 
			this.txtUnit.Location = new System.Drawing.Point(173, 443);
			this.txtUnit.Name = "txtUnit";
			this.txtUnit.Size = new System.Drawing.Size(217, 33);
			this.txtUnit.TabIndex = 7;
			this.txtUnit.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtUnit_KeyDown);
			// 
			// mnuMenu
			// 
			this.mnuMenu.AutoSize = false;
			this.mnuMenu.BackColor = System.Drawing.Color.WhiteSmoke;
			this.mnuMenu.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.mnuMenu.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden;
			this.mnuMenu.ImageScalingSize = new System.Drawing.Size(20, 20);
			this.mnuMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.btnSaveClose});
			this.mnuMenu.LayoutStyle = System.Windows.Forms.ToolStripLayoutStyle.Flow;
			this.mnuMenu.Location = new System.Drawing.Point(0, 0);
			this.mnuMenu.Name = "mnuMenu";
			this.mnuMenu.RenderMode = System.Windows.Forms.ToolStripRenderMode.Professional;
			this.mnuMenu.Size = new System.Drawing.Size(520, 41);
			this.mnuMenu.TabIndex = 227;
			this.mnuMenu.Text = "toolStrip2";
			// 
			// btnSaveClose
			// 
			this.btnSaveClose.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Bold);
			this.btnSaveClose.ForeColor = System.Drawing.Color.Black;
			this.btnSaveClose.Image = ((System.Drawing.Image)(resources.GetObject("btnSaveClose.Image")));
			this.btnSaveClose.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnSaveClose.Name = "btnSaveClose";
			this.btnSaveClose.Size = new System.Drawing.Size(94, 41);
			this.btnSaveClose.Tag = "frmProduct_AddProductH";
			this.btnSaveClose.Text = "Cất && Đóng";
			this.btnSaveClose.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnSaveClose.Click += new System.EventHandler(this.btnSaveClose_Click);
			// 
			// txtMahang
			// 
			this.txtMahang.Enabled = false;
			this.txtMahang.Location = new System.Drawing.Point(173, 103);
			this.txtMahang.Name = "txtMahang";
			this.txtMahang.Size = new System.Drawing.Size(343, 33);
			this.txtMahang.TabIndex = 0;
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(11, 106);
			this.label8.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(87, 26);
			this.label8.TabIndex = 228;
			this.label8.Text = "Mã hàng";
			// 
			// checkBox1
			// 
			this.checkBox1.AutoSize = true;
			this.checkBox1.Location = new System.Drawing.Point(173, 8);
			this.checkBox1.Name = "checkBox1";
			this.checkBox1.Size = new System.Drawing.Size(98, 30);
			this.checkBox1.TabIndex = 229;
			this.checkBox1.Text = "Hiển thị";
			this.checkBox1.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.checkBox1.UseVisualStyleBackColor = true;
			this.checkBox1.CheckedChanged += new System.EventHandler(this.checkBox1_CheckedChanged);
			// 
			// txtCreateby
			// 
			this.txtCreateby.Enabled = false;
			this.txtCreateby.Location = new System.Drawing.Point(173, 55);
			this.txtCreateby.Name = "txtCreateby";
			this.txtCreateby.Size = new System.Drawing.Size(343, 33);
			this.txtCreateby.TabIndex = 230;
			// 
			// lbl
			// 
			this.lbl.AutoSize = true;
			this.lbl.Location = new System.Drawing.Point(11, 58);
			this.lbl.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
			this.lbl.Name = "lbl";
			this.lbl.Size = new System.Drawing.Size(99, 26);
			this.lbl.TabIndex = 231;
			this.lbl.Text = "Người tạo";
			// 
			// menuStrip1
			// 
			this.menuStrip1.Dock = System.Windows.Forms.DockStyle.None;
			this.menuStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
			this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.cấtToolStripMenuItem});
			this.menuStrip1.Location = new System.Drawing.Point(393, 443);
			this.menuStrip1.Name = "menuStrip1";
			this.menuStrip1.Padding = new System.Windows.Forms.Padding(4, 2, 0, 2);
			this.menuStrip1.Size = new System.Drawing.Size(43, 24);
			this.menuStrip1.TabIndex = 232;
			this.menuStrip1.Text = "menuStrip1";
			this.menuStrip1.Visible = false;
			// 
			// cấtToolStripMenuItem
			// 
			this.cấtToolStripMenuItem.Name = "cấtToolStripMenuItem";
			this.cấtToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.X)));
			this.cấtToolStripMenuItem.Size = new System.Drawing.Size(37, 20);
			this.cấtToolStripMenuItem.Text = "Cất";
			this.cấtToolStripMenuItem.Click += new System.EventHandler(this.cấtToolStripMenuItem_Click);
			// 
			// txtY
			// 
			this.txtY.Location = new System.Drawing.Point(173, 540);
			this.txtY.Name = "txtY";
			this.txtY.Size = new System.Drawing.Size(217, 33);
			this.txtY.TabIndex = 236;
			// 
			// txtX
			// 
			this.txtX.Location = new System.Drawing.Point(173, 492);
			this.txtX.Name = "txtX";
			this.txtX.Size = new System.Drawing.Size(217, 33);
			this.txtX.TabIndex = 235;
			this.txtX.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtX_KeyDown);
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.Location = new System.Drawing.Point(11, 543);
			this.label9.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(22, 26);
			this.label9.TabIndex = 234;
			this.label9.Text = "Y";
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.Location = new System.Drawing.Point(11, 495);
			this.label10.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(23, 26);
			this.label10.TabIndex = 233;
			this.label10.Text = "X";
			// 
			// frmHypMoldDetail
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(11F, 26F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(520, 585);
			this.Controls.Add(this.txtY);
			this.Controls.Add(this.txtX);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.label10);
			this.Controls.Add(this.menuStrip1);
			this.Controls.Add(this.txtCreateby);
			this.Controls.Add(this.lbl);
			this.Controls.Add(this.checkBox1);
			this.Controls.Add(this.txtMahang);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.mnuMenu);
			this.Controls.Add(this.txtUnit);
			this.Controls.Add(this.txtMax);
			this.Controls.Add(this.txtMin);
			this.Controls.Add(this.txtStandardValue);
			this.Controls.Add(this.txtStt);
			this.Controls.Add(this.txtCheckingName);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.cboType);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Font = new System.Drawing.Font("Calibri", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.Margin = new System.Windows.Forms.Padding(6);
			this.Name = "frmHypMoldDetail";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
			this.Text = "HypMoldDetail";
			this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.frmHypMoldDetail_FormClosed);
			this.Load += new System.EventHandler(this.frmHypMoldDetail_Load);
			this.mnuMenu.ResumeLayout(false);
			this.mnuMenu.PerformLayout();
			this.menuStrip1.ResumeLayout(false);
			this.menuStrip1.PerformLayout();
			this.ResumeLayout(false);
			this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.ComboBox cboType;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox txtCheckingName;
        private System.Windows.Forms.TextBox txtStt;
        private System.Windows.Forms.TextBox txtStandardValue;
        private System.Windows.Forms.TextBox txtMin;
        private System.Windows.Forms.TextBox txtMax;
        private System.Windows.Forms.TextBox txtUnit;
        private System.Windows.Forms.ToolStrip mnuMenu;
        private System.Windows.Forms.ToolStripButton btnSaveClose;
        private System.Windows.Forms.TextBox txtMahang;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.CheckBox checkBox1;
        private System.Windows.Forms.TextBox txtCreateby;
        private System.Windows.Forms.Label lbl;
		private System.Windows.Forms.MenuStrip menuStrip1;
		private System.Windows.Forms.ToolStripMenuItem cấtToolStripMenuItem;
		private System.Windows.Forms.TextBox txtY;
		private System.Windows.Forms.TextBox txtX;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.Label label10;
	}
}