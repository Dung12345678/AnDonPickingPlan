namespace BMS
{
    partial class frmWorkingDetail
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
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmWorkingDetail));
			this.mnuMenu = new System.Windows.Forms.ToolStrip();
			this.btnSave = new System.Windows.Forms.ToolStripButton();
			this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
			this.btnSaveNew = new System.Windows.Forms.ToolStripButton();
			this.cboGroup = new DevExpress.XtraEditors.LookUpEdit();
			this.labelControl1 = new DevExpress.XtraEditors.LabelControl();
			this.labelControl2 = new DevExpress.XtraEditors.LabelControl();
			this.cboStep = new DevExpress.XtraEditors.LookUpEdit();
			this.txtSortOrder = new System.Windows.Forms.NumericUpDown();
			this.labelControl3 = new DevExpress.XtraEditors.LabelControl();
			this.txtName = new System.Windows.Forms.TextBox();
			this.labelControl4 = new DevExpress.XtraEditors.LabelControl();
			this.labelControl5 = new DevExpress.XtraEditors.LabelControl();
			this.labelControl6 = new DevExpress.XtraEditors.LabelControl();
			this.txtPeriodValue = new System.Windows.Forms.TextBox();
			this.labelControl7 = new DevExpress.XtraEditors.LabelControl();
			this.labelControl8 = new DevExpress.XtraEditors.LabelControl();
			this.cboValueType = new System.Windows.Forms.ComboBox();
			this.txtMin = new System.Windows.Forms.NumericUpDown();
			this.txtMax = new System.Windows.Forms.NumericUpDown();
			this.chkAutoValue = new System.Windows.Forms.CheckBox();
			this.cboComPort = new System.Windows.Forms.ComboBox();
			this.txtUnit = new System.Windows.Forms.TextBox();
			this.labelControl9 = new DevExpress.XtraEditors.LabelControl();
			this.chkIP = new System.Windows.Forms.CheckBox();
			this.txtIP = new System.Windows.Forms.TextBox();
			this.label2 = new System.Windows.Forms.Label();
			this.txtPort = new System.Windows.Forms.TextBox();
			this.label1 = new System.Windows.Forms.Label();
			this.cboCheckValueType = new System.Windows.Forms.ComboBox();
			this.labelControl10 = new DevExpress.XtraEditors.LabelControl();
			this.panel1 = new System.Windows.Forms.Panel();
			this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
			this.saveToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.saveNewToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.closeToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.mnuMenu.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cboGroup.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cboStep.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSortOrder)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMin)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMax)).BeginInit();
			this.panel1.SuspendLayout();
			this.contextMenuStrip1.SuspendLayout();
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
            this.btnSaveNew});
			this.mnuMenu.LayoutStyle = System.Windows.Forms.ToolStripLayoutStyle.Flow;
			this.mnuMenu.Location = new System.Drawing.Point(0, 0);
			this.mnuMenu.Name = "mnuMenu";
			this.mnuMenu.RenderMode = System.Windows.Forms.ToolStripRenderMode.Professional;
			this.mnuMenu.Size = new System.Drawing.Size(554, 42);
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
			this.btnSave.Tag = "frmWorking_AddWorking";
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
			this.btnSaveNew.Tag = "frmWorking_AddWorking";
			this.btnSaveNew.Text = "Cất && Thêm mới";
			this.btnSaveNew.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnSaveNew.Click += new System.EventHandler(this.btnSaveNew_Click);
			// 
			// cboGroup
			// 
			this.cboGroup.Location = new System.Drawing.Point(156, 56);
			this.cboGroup.Name = "cboGroup";
			this.cboGroup.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cboGroup.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
            new DevExpress.XtraEditors.Controls.LookUpColumnInfo("ID", "ID", 20, DevExpress.Utils.FormatType.None, "", false, DevExpress.Utils.HorzAlignment.Default),
            new DevExpress.XtraEditors.Controls.LookUpColumnInfo("ProductGroupCode", 10, "Mã"),
            new DevExpress.XtraEditors.Controls.LookUpColumnInfo("ProductGroupName", "Tên")});
			this.cboGroup.Properties.NullText = "";
			this.cboGroup.Size = new System.Drawing.Size(381, 20);
			this.cboGroup.TabIndex = 8;
			this.cboGroup.EditValueChanged += new System.EventHandler(this.cboGroup_EditValueChanged);
			// 
			// labelControl1
			// 
			this.labelControl1.Location = new System.Drawing.Point(73, 59);
			this.labelControl1.Name = "labelControl1";
			this.labelControl1.Size = new System.Drawing.Size(76, 13);
			this.labelControl1.TabIndex = 29;
			this.labelControl1.Text = "Nhóm sản phẩm";
			// 
			// labelControl2
			// 
			this.labelControl2.Location = new System.Drawing.Point(97, 85);
			this.labelControl2.Name = "labelControl2";
			this.labelControl2.Size = new System.Drawing.Size(52, 13);
			this.labelControl2.TabIndex = 29;
			this.labelControl2.Text = "Công đoạn";
			// 
			// cboStep
			// 
			this.cboStep.Location = new System.Drawing.Point(156, 82);
			this.cboStep.Name = "cboStep";
			this.cboStep.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cboStep.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
            new DevExpress.XtraEditors.Controls.LookUpColumnInfo("ID", "ID", 20, DevExpress.Utils.FormatType.None, "", false, DevExpress.Utils.HorzAlignment.Default),
            new DevExpress.XtraEditors.Controls.LookUpColumnInfo("SortOrder", "Công đoạn", 10, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Center),
            new DevExpress.XtraEditors.Controls.LookUpColumnInfo("WorkingStepCode", 10, "Mã"),
            new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Description", 30, "Tên")});
			this.cboStep.Properties.NullText = "";
			this.cboStep.Size = new System.Drawing.Size(381, 20);
			this.cboStep.TabIndex = 9;
			// 
			// txtSortOrder
			// 
			this.txtSortOrder.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtSortOrder.Location = new System.Drawing.Point(156, 169);
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
			this.txtSortOrder.Size = new System.Drawing.Size(203, 26);
			this.txtSortOrder.TabIndex = 1;
			this.txtSortOrder.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
			// 
			// labelControl3
			// 
			this.labelControl3.Location = new System.Drawing.Point(116, 176);
			this.labelControl3.Name = "labelControl3";
			this.labelControl3.Size = new System.Drawing.Size(33, 13);
			this.labelControl3.TabIndex = 29;
			this.labelControl3.Text = "Thứ tự";
			// 
			// txtName
			// 
			this.txtName.Location = new System.Drawing.Point(156, 108);
			this.txtName.Multiline = true;
			this.txtName.Name = "txtName";
			this.txtName.Size = new System.Drawing.Size(381, 53);
			this.txtName.TabIndex = 0;
			// 
			// labelControl4
			// 
			this.labelControl4.Location = new System.Drawing.Point(122, 118);
			this.labelControl4.Name = "labelControl4";
			this.labelControl4.Size = new System.Drawing.Size(27, 13);
			this.labelControl4.TabIndex = 29;
			this.labelControl4.Text = "Mô tả";
			// 
			// labelControl5
			// 
			this.labelControl5.Location = new System.Drawing.Point(88, 207);
			this.labelControl5.Name = "labelControl5";
			this.labelControl5.Size = new System.Drawing.Size(61, 13);
			this.labelControl5.TabIndex = 29;
			this.labelControl5.Text = "Kiểu kiểm tra";
			// 
			// labelControl6
			// 
			this.labelControl6.Location = new System.Drawing.Point(68, 363);
			this.labelControl6.Name = "labelControl6";
			this.labelControl6.Size = new System.Drawing.Size(81, 13);
			this.labelControl6.TabIndex = 29;
			this.labelControl6.Text = "Giá trị tiêu chuẩn";
			// 
			// txtPeriodValue
			// 
			this.txtPeriodValue.Location = new System.Drawing.Point(156, 360);
			this.txtPeriodValue.Name = "txtPeriodValue";
			this.txtPeriodValue.Size = new System.Drawing.Size(381, 20);
			this.txtPeriodValue.TabIndex = 7;
			// 
			// labelControl7
			// 
			this.labelControl7.Location = new System.Drawing.Point(22, 298);
			this.labelControl7.Name = "labelControl7";
			this.labelControl7.Size = new System.Drawing.Size(127, 13);
			this.labelControl7.TabIndex = 29;
			this.labelControl7.Text = "Giá trị tiêu chuẩn nhỏ nhất";
			// 
			// labelControl8
			// 
			this.labelControl8.Location = new System.Drawing.Point(26, 332);
			this.labelControl8.Name = "labelControl8";
			this.labelControl8.Size = new System.Drawing.Size(123, 13);
			this.labelControl8.TabIndex = 29;
			this.labelControl8.Text = "Giá trị tiêu chuẩn lớn nhất";
			// 
			// cboValueType
			// 
			this.cboValueType.FormattingEnabled = true;
			this.cboValueType.Items.AddRange(new object[] {
            "Check mark",
            "Giá trị"});
			this.cboValueType.Location = new System.Drawing.Point(156, 203);
			this.cboValueType.Name = "cboValueType";
			this.cboValueType.Size = new System.Drawing.Size(203, 21);
			this.cboValueType.TabIndex = 2;
			// 
			// txtMin
			// 
			this.txtMin.DecimalPlaces = 3;
			this.txtMin.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtMin.Location = new System.Drawing.Point(156, 292);
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
			this.txtMin.TabIndex = 5;
			this.txtMin.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
			this.txtMin.ValueChanged += new System.EventHandler(this.txtMin_ValueChanged);
			// 
			// txtMax
			// 
			this.txtMax.DecimalPlaces = 3;
			this.txtMax.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtMax.Location = new System.Drawing.Point(156, 326);
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
			// chkAutoValue
			// 
			this.chkAutoValue.AutoSize = true;
			this.chkAutoValue.Location = new System.Drawing.Point(14, 36);
			this.chkAutoValue.Name = "chkAutoValue";
			this.chkAutoValue.Size = new System.Drawing.Size(179, 17);
			this.chkAutoValue.TabIndex = 35;
			this.chkAutoValue.Text = "Lấy dữ liệu tự động từ cổng com";
			this.chkAutoValue.UseVisualStyleBackColor = true;
			this.chkAutoValue.CheckedChanged += new System.EventHandler(this.chkAutoValue_CheckedChanged);
			// 
			// cboComPort
			// 
			this.cboComPort.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.cboComPort.FormattingEnabled = true;
			this.cboComPort.Items.AddRange(new object[] {
            "COM1",
            "COM2",
            "COM3",
            "COM4",
            "COM5",
            "COM6",
            "COM7",
            "COM8",
            "COM9",
            "COM10"});
			this.cboComPort.Location = new System.Drawing.Point(225, 29);
			this.cboComPort.Name = "cboComPort";
			this.cboComPort.Size = new System.Drawing.Size(121, 28);
			this.cboComPort.TabIndex = 36;
			// 
			// txtUnit
			// 
			this.txtUnit.Location = new System.Drawing.Point(156, 262);
			this.txtUnit.Name = "txtUnit";
			this.txtUnit.Size = new System.Drawing.Size(203, 20);
			this.txtUnit.TabIndex = 4;
			// 
			// labelControl9
			// 
			this.labelControl9.Location = new System.Drawing.Point(118, 265);
			this.labelControl9.Name = "labelControl9";
			this.labelControl9.Size = new System.Drawing.Size(31, 13);
			this.labelControl9.TabIndex = 37;
			this.labelControl9.Text = "Đơn vị";
			// 
			// chkIP
			// 
			this.chkIP.AutoSize = true;
			this.chkIP.Location = new System.Drawing.Point(14, 65);
			this.chkIP.Name = "chkIP";
			this.chkIP.Size = new System.Drawing.Size(220, 17);
			this.chkIP.TabIndex = 43;
			this.chkIP.Text = "Lấy dữ liệu tự động từ truyền thông mạng";
			this.chkIP.UseVisualStyleBackColor = true;
			this.chkIP.CheckedChanged += new System.EventHandler(this.chkIP_CheckedChanged);
			// 
			// txtIP
			// 
			this.txtIP.Location = new System.Drawing.Point(77, 85);
			this.txtIP.Name = "txtIP";
			this.txtIP.Size = new System.Drawing.Size(90, 20);
			this.txtIP.TabIndex = 44;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(280, 87);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(32, 13);
			this.label2.TabIndex = 47;
			this.label2.Text = "Cổng";
			// 
			// txtPort
			// 
			this.txtPort.Location = new System.Drawing.Point(318, 84);
			this.txtPort.Name = "txtPort";
			this.txtPort.Size = new System.Drawing.Size(76, 20);
			this.txtPort.TabIndex = 45;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(18, 88);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(53, 13);
			this.label1.TabIndex = 46;
			this.label1.Text = "Địa chỉ IP";
			// 
			// cboCheckValueType
			// 
			this.cboCheckValueType.FormattingEnabled = true;
			this.cboCheckValueType.Items.AddRange(new object[] {
            "--Chọn một loại--",
            "Số",
            "Ký tự"});
			this.cboCheckValueType.Location = new System.Drawing.Point(156, 232);
			this.cboCheckValueType.Name = "cboCheckValueType";
			this.cboCheckValueType.Size = new System.Drawing.Size(203, 21);
			this.cboCheckValueType.TabIndex = 3;
			// 
			// labelControl10
			// 
			this.labelControl10.Location = new System.Drawing.Point(50, 235);
			this.labelControl10.Name = "labelControl10";
			this.labelControl10.Size = new System.Drawing.Size(99, 13);
			this.labelControl10.TabIndex = 48;
			this.labelControl10.Text = "Loại giá trị kiểm tra *";
			// 
			// panel1
			// 
			this.panel1.Controls.Add(this.cboComPort);
			this.panel1.Controls.Add(this.chkAutoValue);
			this.panel1.Controls.Add(this.label1);
			this.panel1.Controls.Add(this.chkIP);
			this.panel1.Controls.Add(this.txtPort);
			this.panel1.Controls.Add(this.txtIP);
			this.panel1.Controls.Add(this.label2);
			this.panel1.Location = new System.Drawing.Point(399, 203);
			this.panel1.Name = "panel1";
			this.panel1.Size = new System.Drawing.Size(112, 129);
			this.panel1.TabIndex = 49;
			this.panel1.Visible = false;
			// 
			// contextMenuStrip1
			// 
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
			this.closeToolStripMenuItem.Text = "Close";
			this.closeToolStripMenuItem.Click += new System.EventHandler(this.closeToolStripMenuItem_Click);
			// 
			// frmWorkingDetail
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(554, 393);
			this.ContextMenuStrip = this.contextMenuStrip1;
			this.Controls.Add(this.panel1);
			this.Controls.Add(this.cboCheckValueType);
			this.Controls.Add(this.labelControl10);
			this.Controls.Add(this.txtUnit);
			this.Controls.Add(this.labelControl9);
			this.Controls.Add(this.cboValueType);
			this.Controls.Add(this.txtPeriodValue);
			this.Controls.Add(this.txtName);
			this.Controls.Add(this.labelControl8);
			this.Controls.Add(this.txtMax);
			this.Controls.Add(this.txtMin);
			this.Controls.Add(this.txtSortOrder);
			this.Controls.Add(this.labelControl7);
			this.Controls.Add(this.cboStep);
			this.Controls.Add(this.labelControl6);
			this.Controls.Add(this.labelControl5);
			this.Controls.Add(this.labelControl4);
			this.Controls.Add(this.labelControl3);
			this.Controls.Add(this.labelControl2);
			this.Controls.Add(this.cboGroup);
			this.Controls.Add(this.labelControl1);
			this.Controls.Add(this.mnuMenu);
			this.Name = "frmWorkingDetail";
			this.Text = "CHI TIẾT MỤC KIỂM TRA";
			this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmWorkingDetail_FormClosing);
			this.Load += new System.EventHandler(this.frmWorkingDetail_Load);
			this.mnuMenu.ResumeLayout(false);
			this.mnuMenu.PerformLayout();
			((System.ComponentModel.ISupportInitialize)(this.cboGroup.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cboStep.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSortOrder)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMin)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMax)).EndInit();
			this.panel1.ResumeLayout(false);
			this.panel1.PerformLayout();
			this.contextMenuStrip1.ResumeLayout(false);
			this.ResumeLayout(false);
			this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ToolStrip mnuMenu;
        private System.Windows.Forms.ToolStripButton btnSave;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
        private DevExpress.XtraEditors.LookUpEdit cboGroup;
        private DevExpress.XtraEditors.LabelControl labelControl1;
        private DevExpress.XtraEditors.LabelControl labelControl2;
        private DevExpress.XtraEditors.LookUpEdit cboStep;
        private System.Windows.Forms.NumericUpDown txtSortOrder;
        private DevExpress.XtraEditors.LabelControl labelControl3;
        private System.Windows.Forms.TextBox txtName;
        private DevExpress.XtraEditors.LabelControl labelControl4;
        private DevExpress.XtraEditors.LabelControl labelControl5;
        private DevExpress.XtraEditors.LabelControl labelControl6;
        private System.Windows.Forms.TextBox txtPeriodValue;
        private DevExpress.XtraEditors.LabelControl labelControl7;
        private DevExpress.XtraEditors.LabelControl labelControl8;
        private System.Windows.Forms.ComboBox cboValueType;
        private System.Windows.Forms.NumericUpDown txtMin;
        private System.Windows.Forms.NumericUpDown txtMax;
        private System.Windows.Forms.CheckBox chkAutoValue;
        private System.Windows.Forms.ComboBox cboComPort;
        private System.Windows.Forms.ToolStripButton btnSaveNew;
        private System.Windows.Forms.TextBox txtUnit;
        private DevExpress.XtraEditors.LabelControl labelControl9;
        private System.Windows.Forms.CheckBox chkIP;
        private System.Windows.Forms.TextBox txtIP;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtPort;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox cboCheckValueType;
        private DevExpress.XtraEditors.LabelControl labelControl10;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem saveToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem saveNewToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem closeToolStripMenuItem;
    }
}