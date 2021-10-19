namespace BMS
{
    partial class frmShelfDetail
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
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmShelfDetail));
			this.txtTenGiaDo = new System.Windows.Forms.TextBox();
			this.labelControl9 = new DevExpress.XtraEditors.LabelControl();
			this.txtGiaDo = new System.Windows.Forms.TextBox();
			this.labelControl4 = new DevExpress.XtraEditors.LabelControl();
			this.cboGroupnhom = new DevExpress.XtraEditors.LookUpEdit();
			this.mnuMenu = new System.Windows.Forms.ToolStrip();
			this.btnSave = new System.Windows.Forms.ToolStripButton();
			this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
			this.btnSaveNew = new System.Windows.Forms.ToolStripButton();
			this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
			this.saveToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.SaveNewToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.txtSTT = new System.Windows.Forms.TextBox();
			this.labelControl6 = new DevExpress.XtraEditors.LabelControl();
			this.labelControl1 = new DevExpress.XtraEditors.LabelControl();
			((System.ComponentModel.ISupportInitialize)(this.cboGroupnhom.Properties)).BeginInit();
			this.mnuMenu.SuspendLayout();
			this.contextMenuStrip1.SuspendLayout();
			this.SuspendLayout();
			// 
			// txtTenGiaDo
			// 
			this.txtTenGiaDo.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtTenGiaDo.Location = new System.Drawing.Point(77, 183);
			this.txtTenGiaDo.Name = "txtTenGiaDo";
			this.txtTenGiaDo.Size = new System.Drawing.Size(229, 21);
			this.txtTenGiaDo.TabIndex = 3;
			this.txtTenGiaDo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtTenGiaDo_KeyDown);
			// 
			// labelControl9
			// 
			this.labelControl9.Location = new System.Drawing.Point(23, 187);
			this.labelControl9.Name = "labelControl9";
			this.labelControl9.Size = new System.Drawing.Size(50, 13);
			this.labelControl9.TabIndex = 186;
			this.labelControl9.Text = "Tên giá đỡ";
			// 
			// txtGiaDo
			// 
			this.txtGiaDo.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtGiaDo.Location = new System.Drawing.Point(77, 140);
			this.txtGiaDo.Name = "txtGiaDo";
			this.txtGiaDo.Size = new System.Drawing.Size(229, 21);
			this.txtGiaDo.TabIndex = 2;
			this.txtGiaDo.Tag = "";
			this.txtGiaDo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtGiaDo_KeyDown);
			// 
			// labelControl4
			// 
			this.labelControl4.Location = new System.Drawing.Point(23, 144);
			this.labelControl4.Name = "labelControl4";
			this.labelControl4.Size = new System.Drawing.Size(46, 13);
			this.labelControl4.TabIndex = 183;
			this.labelControl4.Text = "Mã giá đỡ";
			// 
			// cboGroupnhom
			// 
			this.cboGroupnhom.Location = new System.Drawing.Point(77, 100);
			this.cboGroupnhom.Name = "cboGroupnhom";
			this.cboGroupnhom.Properties.Appearance.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.cboGroupnhom.Properties.Appearance.Options.UseFont = true;
			this.cboGroupnhom.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cboGroupnhom.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
            new DevExpress.XtraEditors.Controls.LookUpColumnInfo("ID", "ID", 20, DevExpress.Utils.FormatType.None, "", false, DevExpress.Utils.HorzAlignment.Default),
            new DevExpress.XtraEditors.Controls.LookUpColumnInfo("SequenceCode", "Mã dãy"),
            new DevExpress.XtraEditors.Controls.LookUpColumnInfo("SequenceName", 40, "Tên dãy")});
			this.cboGroupnhom.Properties.NullText = "";
			this.cboGroupnhom.Size = new System.Drawing.Size(229, 22);
			this.cboGroupnhom.TabIndex = 100;
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
			this.mnuMenu.Size = new System.Drawing.Size(318, 42);
			this.mnuMenu.TabIndex = 196;
			this.mnuMenu.Text = "toolStrip2";
			// 
			// btnSave
			// 
			this.btnSave.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnSave.Image = ((System.Drawing.Image)(resources.GetObject("btnSave.Image")));
			this.btnSave.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnSave.Name = "btnSave";
			this.btnSave.Size = new System.Drawing.Size(75, 37);
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
			this.btnSaveNew.Text = "Cất && Thêm mới";
			this.btnSaveNew.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnSaveNew.Click += new System.EventHandler(this.btnSaveNew_Click);
			// 
			// contextMenuStrip1
			// 
			this.contextMenuStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
			this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.saveToolStripMenuItem,
            this.SaveNewToolStripMenuItem});
			this.contextMenuStrip1.Name = "contextMenuStrip1";
			this.contextMenuStrip1.Size = new System.Drawing.Size(163, 48);
			// 
			// saveToolStripMenuItem
			// 
			this.saveToolStripMenuItem.Name = "saveToolStripMenuItem";
			this.saveToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.X)));
			this.saveToolStripMenuItem.Size = new System.Drawing.Size(162, 22);
			this.saveToolStripMenuItem.Text = "Save";
			this.saveToolStripMenuItem.Visible = false;
			this.saveToolStripMenuItem.Click += new System.EventHandler(this.saveToolStripMenuItem_Click);
			// 
			// SaveNewToolStripMenuItem
			// 
			this.SaveNewToolStripMenuItem.Name = "SaveNewToolStripMenuItem";
			this.SaveNewToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.S)));
			this.SaveNewToolStripMenuItem.Size = new System.Drawing.Size(162, 22);
			this.SaveNewToolStripMenuItem.Text = "SaveNew";
			this.SaveNewToolStripMenuItem.Click += new System.EventHandler(this.SaveNewToolStripMenuItem_Click);
			// 
			// txtSTT
			// 
			this.txtSTT.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtSTT.Location = new System.Drawing.Point(77, 59);
			this.txtSTT.Name = "txtSTT";
			this.txtSTT.Size = new System.Drawing.Size(229, 21);
			this.txtSTT.TabIndex = 0;
			this.txtSTT.Tag = "";
			this.txtSTT.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSTT_KeyDown);
			// 
			// labelControl6
			// 
			this.labelControl6.Location = new System.Drawing.Point(23, 63);
			this.labelControl6.Name = "labelControl6";
			this.labelControl6.Size = new System.Drawing.Size(46, 13);
			this.labelControl6.TabIndex = 209;
			this.labelControl6.Text = "Số thứ tự";
			// 
			// labelControl1
			// 
			this.labelControl1.Location = new System.Drawing.Point(23, 105);
			this.labelControl1.Name = "labelControl1";
			this.labelControl1.Size = new System.Drawing.Size(35, 13);
			this.labelControl1.TabIndex = 214;
			this.labelControl1.Text = "Mã dãy";
			// 
			// frmShelfDetail
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(318, 223);
			this.ContextMenuStrip = this.contextMenuStrip1;
			this.Controls.Add(this.labelControl1);
			this.Controls.Add(this.labelControl6);
			this.Controls.Add(this.txtSTT);
			this.Controls.Add(this.mnuMenu);
			this.Controls.Add(this.txtTenGiaDo);
			this.Controls.Add(this.labelControl9);
			this.Controls.Add(this.txtGiaDo);
			this.Controls.Add(this.labelControl4);
			this.Controls.Add(this.cboGroupnhom);
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.Margin = new System.Windows.Forms.Padding(4);
			this.Name = "frmShelfDetail";
			this.Text = "THÔNG TIN GIÁ ĐỠ";
			this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmShelfDetail_FormClosing);
			this.Load += new System.EventHandler(this.frmShelfDetail_Load);
			((System.ComponentModel.ISupportInitialize)(this.cboGroupnhom.Properties)).EndInit();
			this.mnuMenu.ResumeLayout(false);
			this.mnuMenu.PerformLayout();
			this.contextMenuStrip1.ResumeLayout(false);
			this.ResumeLayout(false);
			this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.TextBox txtTenGiaDo;
        private DevExpress.XtraEditors.LabelControl labelControl9;
        private System.Windows.Forms.TextBox txtGiaDo;
        private DevExpress.XtraEditors.LabelControl labelControl4;
        private DevExpress.XtraEditors.LookUpEdit cboGroupnhom;
        private System.Windows.Forms.ToolStrip mnuMenu;
        private System.Windows.Forms.ToolStripButton btnSave;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
        private System.Windows.Forms.ToolStripButton btnSaveNew;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem saveToolStripMenuItem;
        private System.Windows.Forms.TextBox txtSTT;
        private DevExpress.XtraEditors.LabelControl labelControl6;
        private DevExpress.XtraEditors.LabelControl labelControl1;
        private System.Windows.Forms.ToolStripMenuItem SaveNewToolStripMenuItem;
    }
}