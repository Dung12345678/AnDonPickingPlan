namespace BMS
{
    partial class frmGearWorkingDetail
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmGearWorkingDetail));
            this.mnuMenu = new System.Windows.Forms.ToolStrip();
            this.btnSave = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.btnSaveNew = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.btnClose = new System.Windows.Forms.ToolStripButton();
            this.txtUnit = new System.Windows.Forms.TextBox();
            this.labelControl9 = new DevExpress.XtraEditors.LabelControl();
            this.cboCheckValueType = new System.Windows.Forms.ComboBox();
            this.cboValueType = new System.Windows.Forms.ComboBox();
            this.txtPeriodValue = new System.Windows.Forms.TextBox();
            this.labelControl8 = new DevExpress.XtraEditors.LabelControl();
            this.txtMax = new System.Windows.Forms.NumericUpDown();
            this.txtMin = new System.Windows.Forms.NumericUpDown();
            this.txtSortOrder = new System.Windows.Forms.NumericUpDown();
            this.labelControl7 = new DevExpress.XtraEditors.LabelControl();
            this.labelControl1 = new DevExpress.XtraEditors.LabelControl();
            this.labelControl6 = new DevExpress.XtraEditors.LabelControl();
            this.labelControl5 = new DevExpress.XtraEditors.LabelControl();
            this.labelControl3 = new DevExpress.XtraEditors.LabelControl();
            this.txtName = new System.Windows.Forms.TextBox();
            this.labelControl4 = new DevExpress.XtraEditors.LabelControl();
            this.labelControl2 = new DevExpress.XtraEditors.LabelControl();
            this.numDefautValue = new System.Windows.Forms.NumericUpDown();
            this.labelControl10 = new DevExpress.XtraEditors.LabelControl();
            this.txtTanSuat = new System.Windows.Forms.TextBox();
            this.labelControl11 = new DevExpress.XtraEditors.LabelControl();
            this.mnuMenu.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.txtMax)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtMin)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtSortOrder)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numDefautValue)).BeginInit();
            this.SuspendLayout();
            // 
            // mnuMenu
            // 
            this.mnuMenu.AutoSize = false;
            this.mnuMenu.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(234)))), ((int)(((byte)(234)))), ((int)(((byte)(234)))));
            this.mnuMenu.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.mnuMenu.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden;
            this.mnuMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.btnSave,
            this.toolStripSeparator2,
            this.btnSaveNew,
            this.toolStripSeparator1,
            this.btnClose});
            this.mnuMenu.LayoutStyle = System.Windows.Forms.ToolStripLayoutStyle.Flow;
            this.mnuMenu.Location = new System.Drawing.Point(0, 0);
            this.mnuMenu.Name = "mnuMenu";
            this.mnuMenu.RenderMode = System.Windows.Forms.ToolStripRenderMode.Professional;
            this.mnuMenu.Size = new System.Drawing.Size(547, 42);
            this.mnuMenu.TabIndex = 10;
            this.mnuMenu.Text = "toolStrip2";
            // 
            // btnSave
            // 
            this.btnSave.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSave.Image = ((System.Drawing.Image)(resources.GetObject("btnSave.Image")));
            this.btnSave.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 33);
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
            this.btnSaveNew.Size = new System.Drawing.Size(101, 33);
            this.btnSaveNew.Text = "Cất && Thêm mới";
            this.btnSaveNew.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.btnSaveNew.Click += new System.EventHandler(this.btnSaveNew_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.AutoSize = false;
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(6, 40);
            // 
            // btnClose
            // 
            this.btnClose.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnClose.Image = ((System.Drawing.Image)(resources.GetObject("btnClose.Image")));
            this.btnClose.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(44, 33);
            this.btnClose.Text = "Thoát";
            this.btnClose.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // txtUnit
            // 
            this.txtUnit.Location = new System.Drawing.Point(148, 199);
            this.txtUnit.Name = "txtUnit";
            this.txtUnit.Size = new System.Drawing.Size(203, 20);
            this.txtUnit.TabIndex = 4;
            // 
            // labelControl9
            // 
            this.labelControl9.Location = new System.Drawing.Point(101, 203);
            this.labelControl9.Name = "labelControl9";
            this.labelControl9.Size = new System.Drawing.Size(31, 13);
            this.labelControl9.TabIndex = 62;
            this.labelControl9.Text = "Đơn vị";
            // 
            // cboCheckValueType
            // 
            this.cboCheckValueType.FormattingEnabled = true;
            this.cboCheckValueType.Items.AddRange(new object[] {
            "--Chọn một loại--",
            "Số",
            "Ký tự"});
            this.cboCheckValueType.Location = new System.Drawing.Point(148, 166);
            this.cboCheckValueType.Name = "cboCheckValueType";
            this.cboCheckValueType.Size = new System.Drawing.Size(203, 21);
            this.cboCheckValueType.TabIndex = 3;
            // 
            // cboValueType
            // 
            this.cboValueType.FormattingEnabled = true;
            this.cboValueType.Items.AddRange(new object[] {
            "Check mark",
            "Giá trị"});
            this.cboValueType.Location = new System.Drawing.Point(148, 133);
            this.cboValueType.Name = "cboValueType";
            this.cboValueType.Size = new System.Drawing.Size(203, 21);
            this.cboValueType.TabIndex = 2;
            // 
            // txtPeriodValue
            // 
            this.txtPeriodValue.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.txtPeriodValue.Location = new System.Drawing.Point(148, 378);
            this.txtPeriodValue.Name = "txtPeriodValue";
            this.txtPeriodValue.Size = new System.Drawing.Size(381, 26);
            this.txtPeriodValue.TabIndex = 9;
            // 
            // labelControl8
            // 
            this.labelControl8.Location = new System.Drawing.Point(9, 270);
            this.labelControl8.Name = "labelControl8";
            this.labelControl8.Size = new System.Drawing.Size(129, 13);
            this.labelControl8.TabIndex = 60;
            this.labelControl8.Text = "Giá trị tiêu chuẩn lớn nhất*";
            // 
            // txtMax
            // 
            this.txtMax.DecimalPlaces = 3;
            this.txtMax.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtMax.Location = new System.Drawing.Point(148, 263);
            this.txtMax.Maximum = new decimal(new int[] {
            1000000000,
            0,
            0,
            0});
            this.txtMax.Minimum = new decimal(new int[] {
            1000000000,
            0,
            0,
            -2147483648});
            this.txtMax.Name = "txtMax";
            this.txtMax.Size = new System.Drawing.Size(203, 26);
            this.txtMax.TabIndex = 6;
            this.txtMax.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.txtMax.ValueChanged += new System.EventHandler(this.txtMax_ValueChanged);
            // 
            // txtMin
            // 
            this.txtMin.DecimalPlaces = 3;
            this.txtMin.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtMin.Location = new System.Drawing.Point(148, 299);
            this.txtMin.Maximum = new decimal(new int[] {
            1000000000,
            0,
            0,
            0});
            this.txtMin.Minimum = new decimal(new int[] {
            1000000000,
            0,
            0,
            -2147483648});
            this.txtMin.Name = "txtMin";
            this.txtMin.Size = new System.Drawing.Size(203, 26);
            this.txtMin.TabIndex = 7;
            this.txtMin.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.txtMin.ValueChanged += new System.EventHandler(this.txtMin_ValueChanged);
            // 
            // txtSortOrder
            // 
            this.txtSortOrder.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtSortOrder.Location = new System.Drawing.Point(148, 99);
            this.txtSortOrder.Maximum = new decimal(new int[] {
            1000000000,
            0,
            0,
            0});
            this.txtSortOrder.Minimum = new decimal(new int[] {
            1000000000,
            0,
            0,
            -2147483648});
            this.txtSortOrder.Name = "txtSortOrder";
            this.txtSortOrder.Size = new System.Drawing.Size(107, 26);
            this.txtSortOrder.TabIndex = 1;
            this.txtSortOrder.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // labelControl7
            // 
            this.labelControl7.Location = new System.Drawing.Point(5, 306);
            this.labelControl7.Name = "labelControl7";
            this.labelControl7.Size = new System.Drawing.Size(133, 13);
            this.labelControl7.TabIndex = 61;
            this.labelControl7.Text = "Giá trị tiêu chuẩn nhỏ nhất*";
            // 
            // labelControl1
            // 
            this.labelControl1.Location = new System.Drawing.Point(33, 170);
            this.labelControl1.Name = "labelControl1";
            this.labelControl1.Size = new System.Drawing.Size(99, 13);
            this.labelControl1.TabIndex = 57;
            this.labelControl1.Text = "Loại giá trị kiểm tra *";
            // 
            // labelControl6
            // 
            this.labelControl6.Location = new System.Drawing.Point(86, 387);
            this.labelControl6.Name = "labelControl6";
            this.labelControl6.Size = new System.Drawing.Size(51, 13);
            this.labelControl6.TabIndex = 56;
            this.labelControl6.Text = "Dải giá trị*";
            // 
            // labelControl5
            // 
            this.labelControl5.Location = new System.Drawing.Point(62, 138);
            this.labelControl5.Name = "labelControl5";
            this.labelControl5.Size = new System.Drawing.Size(70, 13);
            this.labelControl5.TabIndex = 55;
            this.labelControl5.Text = "Kiểu kiểm tra *";
            // 
            // labelControl3
            // 
            this.labelControl3.Location = new System.Drawing.Point(93, 107);
            this.labelControl3.Name = "labelControl3";
            this.labelControl3.Size = new System.Drawing.Size(39, 13);
            this.labelControl3.TabIndex = 59;
            this.labelControl3.Text = "Thứ tự*";
            // 
            // txtName
            // 
            this.txtName.Location = new System.Drawing.Point(148, 53);
            this.txtName.Multiline = true;
            this.txtName.Name = "txtName";
            this.txtName.Size = new System.Drawing.Size(203, 40);
            this.txtName.TabIndex = 0;
            // 
            // labelControl4
            // 
            this.labelControl4.Location = new System.Drawing.Point(97, 69);
            this.labelControl4.Name = "labelControl4";
            this.labelControl4.Size = new System.Drawing.Size(36, 13);
            this.labelControl4.TabIndex = 64;
            this.labelControl4.Text = "Mô tả *";
            // 
            // labelControl2
            // 
            this.labelControl2.Location = new System.Drawing.Point(51, 236);
            this.labelControl2.Name = "labelControl2";
            this.labelControl2.Size = new System.Drawing.Size(87, 13);
            this.labelControl2.TabIndex = 60;
            this.labelControl2.Text = "Giá trị tiêu chuẩn*";
            // 
            // numDefautValue
            // 
            this.numDefautValue.DecimalPlaces = 3;
            this.numDefautValue.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.numDefautValue.Location = new System.Drawing.Point(148, 228);
            this.numDefautValue.Maximum = new decimal(new int[] {
            100000000,
            0,
            0,
            0});
            this.numDefautValue.Minimum = new decimal(new int[] {
            100000000,
            0,
            0,
            -2147483648});
            this.numDefautValue.Name = "numDefautValue";
            this.numDefautValue.Size = new System.Drawing.Size(203, 26);
            this.numDefautValue.TabIndex = 5;
            this.numDefautValue.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // labelControl10
            // 
            this.labelControl10.Location = new System.Drawing.Point(89, 346);
            this.labelControl10.Name = "labelControl10";
            this.labelControl10.Size = new System.Drawing.Size(48, 13);
            this.labelControl10.TabIndex = 60;
            this.labelControl10.Text = "Tần suất*";
            // 
            // txtTanSuat
            // 
            this.txtTanSuat.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.txtTanSuat.Location = new System.Drawing.Point(148, 337);
            this.txtTanSuat.Name = "txtTanSuat";
            this.txtTanSuat.Size = new System.Drawing.Size(100, 26);
            this.txtTanSuat.TabIndex = 8;
            // 
            // labelControl11
            // 
            this.labelControl11.Appearance.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.labelControl11.Location = new System.Drawing.Point(254, 345);
            this.labelControl11.Name = "labelControl11";
            this.labelControl11.Size = new System.Drawing.Size(89, 13);
            this.labelControl11.TabIndex = 60;
            this.labelControl11.Text = "(vd: 1/5, 1/10,...)";
            // 
            // frmGearWorkingDetail
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(547, 429);
            this.Controls.Add(this.txtTanSuat);
            this.Controls.Add(this.numDefautValue);
            this.Controls.Add(this.txtName);
            this.Controls.Add(this.labelControl4);
            this.Controls.Add(this.txtUnit);
            this.Controls.Add(this.labelControl9);
            this.Controls.Add(this.cboCheckValueType);
            this.Controls.Add(this.cboValueType);
            this.Controls.Add(this.txtPeriodValue);
            this.Controls.Add(this.labelControl10);
            this.Controls.Add(this.labelControl11);
            this.Controls.Add(this.labelControl2);
            this.Controls.Add(this.labelControl8);
            this.Controls.Add(this.txtMax);
            this.Controls.Add(this.txtMin);
            this.Controls.Add(this.txtSortOrder);
            this.Controls.Add(this.labelControl7);
            this.Controls.Add(this.labelControl1);
            this.Controls.Add(this.labelControl6);
            this.Controls.Add(this.labelControl5);
            this.Controls.Add(this.labelControl3);
            this.Controls.Add(this.mnuMenu);
            this.Name = "frmGearWorkingDetail";
            this.Text = "CHI TIẾT MỤC KIỂM TRA";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmGearWorkingDetail_FormClosing);
            this.Load += new System.EventHandler(this.frmGearWorkingDetail_Load);
            this.mnuMenu.ResumeLayout(false);
            this.mnuMenu.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.txtMax)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtMin)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtSortOrder)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numDefautValue)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ToolStrip mnuMenu;
        private System.Windows.Forms.ToolStripButton btnSave;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
        private System.Windows.Forms.ToolStripButton btnSaveNew;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripButton btnClose;
        private System.Windows.Forms.TextBox txtUnit;
        private DevExpress.XtraEditors.LabelControl labelControl9;
        private System.Windows.Forms.ComboBox cboCheckValueType;
        private System.Windows.Forms.ComboBox cboValueType;
        private System.Windows.Forms.TextBox txtPeriodValue;
        private DevExpress.XtraEditors.LabelControl labelControl8;
        private System.Windows.Forms.NumericUpDown txtMax;
        private System.Windows.Forms.NumericUpDown txtMin;
        private System.Windows.Forms.NumericUpDown txtSortOrder;
        private DevExpress.XtraEditors.LabelControl labelControl7;
        private DevExpress.XtraEditors.LabelControl labelControl1;
        private DevExpress.XtraEditors.LabelControl labelControl6;
        private DevExpress.XtraEditors.LabelControl labelControl5;
        private DevExpress.XtraEditors.LabelControl labelControl3;
        private System.Windows.Forms.TextBox txtName;
        private DevExpress.XtraEditors.LabelControl labelControl4;
        private DevExpress.XtraEditors.LabelControl labelControl2;
        private System.Windows.Forms.NumericUpDown numDefautValue;
        private DevExpress.XtraEditors.LabelControl labelControl10;
        private System.Windows.Forms.TextBox txtTanSuat;
        private DevExpress.XtraEditors.LabelControl labelControl11;
    }
}