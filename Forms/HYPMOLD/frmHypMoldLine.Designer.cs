
namespace BMS
{
    partial class frmHypMoldLine
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
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmHypMoldLine));
			this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
			this.panel2 = new System.Windows.Forms.Panel();
			this.lbly = new System.Windows.Forms.Label();
			this.lblx = new System.Windows.Forms.Label();
			this.pdfViewer1 = new DevExpress.XtraPdfViewer.PdfViewer();
			this.panel1 = new System.Windows.Forms.Panel();
			this.tableLayoutPanel2 = new System.Windows.Forms.TableLayoutPanel();
			this.grvKnifeJigWorking = new System.Windows.Forms.DataGridView();
			this.colSTT = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.colWorkingName = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.colConditionDisplay = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.colPeriodValue = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.colUnit = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.colRealvalue = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.colEvaluate = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.colY = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.colX = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.colMinValue = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.colID = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.colMaxValue = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.colDisplay = new System.Windows.Forms.DataGridViewCheckBoxColumn();
			this.colCheckValueType = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.panel3 = new System.Windows.Forms.Panel();
			this.textBox1 = new System.Windows.Forms.TextBox();
			this.mnuMenu = new System.Windows.Forms.ToolStrip();
			this.btnSave = new System.Windows.Forms.ToolStripButton();
			this.btnPath = new System.Windows.Forms.ToolStripButton();
			this.txtDateTime = new System.Windows.Forms.TextBox();
			this.txtMachine = new System.Windows.Forms.TextBox();
			this.cboCongDoan = new System.Windows.Forms.ComboBox();
			this.txtOrder = new System.Windows.Forms.TextBox();
			this.txtMaHang = new System.Windows.Forms.TextBox();
			this.txtWorkerCode = new System.Windows.Forms.TextBox();
			this.button4 = new System.Windows.Forms.Button();
			this.button7 = new System.Windows.Forms.Button();
			this.button6 = new System.Windows.Forms.Button();
			this.button3 = new System.Windows.Forms.Button();
			this.button5 = new System.Windows.Forms.Button();
			this.button2 = new System.Windows.Forms.Button();
			this.button1 = new System.Windows.Forms.Button();
			this.tableLayoutPanel1.SuspendLayout();
			this.panel2.SuspendLayout();
			this.panel1.SuspendLayout();
			this.tableLayoutPanel2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.grvKnifeJigWorking)).BeginInit();
			this.panel3.SuspendLayout();
			this.mnuMenu.SuspendLayout();
			this.SuspendLayout();
			// 
			// tableLayoutPanel1
			// 
			this.tableLayoutPanel1.ColumnCount = 1;
			this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
			this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 20F));
			this.tableLayoutPanel1.Controls.Add(this.panel2, 0, 1);
			this.tableLayoutPanel1.Controls.Add(this.panel1, 0, 0);
			this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
			this.tableLayoutPanel1.Name = "tableLayoutPanel1";
			this.tableLayoutPanel1.RowCount = 2;
			this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 21.54006F));
			this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 78.45994F));
			this.tableLayoutPanel1.Size = new System.Drawing.Size(1280, 961);
			this.tableLayoutPanel1.TabIndex = 0;
			// 
			// panel2
			// 
			this.panel2.Controls.Add(this.lbly);
			this.panel2.Controls.Add(this.lblx);
			this.panel2.Controls.Add(this.pdfViewer1);
			this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
			this.panel2.Location = new System.Drawing.Point(3, 209);
			this.panel2.Name = "panel2";
			this.panel2.Size = new System.Drawing.Size(1274, 749);
			this.panel2.TabIndex = 1;
			// 
			// lbly
			// 
			this.lbly.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lbly.AutoSize = true;
			this.lbly.BackColor = System.Drawing.SystemColors.Control;
			this.lbly.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lbly.Location = new System.Drawing.Point(1216, 3);
			this.lbly.Name = "lbly";
			this.lbly.Size = new System.Drawing.Size(15, 16);
			this.lbly.TabIndex = 1;
			this.lbly.Text = "y";
			// 
			// lblx
			// 
			this.lblx.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblx.AutoSize = true;
			this.lblx.BackColor = System.Drawing.SystemColors.Control;
			this.lblx.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblx.Location = new System.Drawing.Point(1176, 3);
			this.lblx.Name = "lblx";
			this.lblx.Size = new System.Drawing.Size(14, 16);
			this.lblx.TabIndex = 0;
			this.lblx.Text = "x";
			// 
			// pdfViewer1
			// 
			this.pdfViewer1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.pdfViewer1.Cursor = System.Windows.Forms.Cursors.Default;
			this.pdfViewer1.Location = new System.Drawing.Point(0, 3);
			this.pdfViewer1.LookAndFeel.SkinName = "VS2010";
			this.pdfViewer1.LookAndFeel.Style = DevExpress.LookAndFeel.LookAndFeelStyle.UltraFlat;
			this.pdfViewer1.Name = "pdfViewer1";
			this.pdfViewer1.Size = new System.Drawing.Size(1271, 746);
			this.pdfViewer1.TabIndex = 0;
			this.pdfViewer1.ZoomMode = DevExpress.XtraPdfViewer.PdfZoomMode.PageLevel;
			// 
			// panel1
			// 
			this.panel1.Controls.Add(this.tableLayoutPanel2);
			this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.panel1.Location = new System.Drawing.Point(3, 3);
			this.panel1.Name = "panel1";
			this.panel1.Size = new System.Drawing.Size(1274, 200);
			this.panel1.TabIndex = 0;
			// 
			// tableLayoutPanel2
			// 
			this.tableLayoutPanel2.ColumnCount = 1;
			this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
			this.tableLayoutPanel2.Controls.Add(this.grvKnifeJigWorking, 0, 1);
			this.tableLayoutPanel2.Controls.Add(this.panel3, 0, 0);
			this.tableLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Fill;
			this.tableLayoutPanel2.Location = new System.Drawing.Point(0, 0);
			this.tableLayoutPanel2.Name = "tableLayoutPanel2";
			this.tableLayoutPanel2.RowCount = 2;
			this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 32F));
			this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 68F));
			this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
			this.tableLayoutPanel2.Size = new System.Drawing.Size(1274, 200);
			this.tableLayoutPanel2.TabIndex = 0;
			// 
			// grvKnifeJigWorking
			// 
			this.grvKnifeJigWorking.AllowUserToAddRows = false;
			this.grvKnifeJigWorking.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
			this.grvKnifeJigWorking.BackgroundColor = System.Drawing.Color.White;
			this.grvKnifeJigWorking.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.Single;
			dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
			dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
			dataGridViewCellStyle1.Font = new System.Drawing.Font("Calibri", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
			dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
			dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
			dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
			this.grvKnifeJigWorking.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
			this.grvKnifeJigWorking.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
			this.grvKnifeJigWorking.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colSTT,
            this.colWorkingName,
            this.colConditionDisplay,
            this.colPeriodValue,
            this.colUnit,
            this.colRealvalue,
            this.colEvaluate,
            this.colY,
            this.colX,
            this.colMinValue,
            this.colID,
            this.colMaxValue,
            this.colDisplay,
            this.colCheckValueType});
			this.grvKnifeJigWorking.Dock = System.Windows.Forms.DockStyle.Bottom;
			this.grvKnifeJigWorking.Location = new System.Drawing.Point(3, 67);
			this.grvKnifeJigWorking.Name = "grvKnifeJigWorking";
			this.grvKnifeJigWorking.RowHeadersWidth = 10;
			this.grvKnifeJigWorking.RowTemplate.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
			this.grvKnifeJigWorking.RowTemplate.DefaultCellStyle.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.grvKnifeJigWorking.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
			this.grvKnifeJigWorking.Size = new System.Drawing.Size(1268, 130);
			this.grvKnifeJigWorking.TabIndex = 74;
			this.grvKnifeJigWorking.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.grvKnifeJigWorking_CellClick);
			this.grvKnifeJigWorking.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.grvKnifeJigWorking_CellValueChanged);
			this.grvKnifeJigWorking.KeyDown += new System.Windows.Forms.KeyEventHandler(this.grvKnifeJigWorking_KeyDown);
			// 
			// colSTT
			// 
			this.colSTT.DataPropertyName = "SortOrder";
			this.colSTT.FillWeight = 30.38773F;
			this.colSTT.HeaderText = "STT";
			this.colSTT.Name = "colSTT";
			this.colSTT.ReadOnly = true;
			// 
			// colWorkingName
			// 
			this.colWorkingName.DataPropertyName = "WorkingName";
			this.colWorkingName.FillWeight = 127.3768F;
			this.colWorkingName.HeaderText = "Mục kiểm tra";
			this.colWorkingName.Name = "colWorkingName";
			this.colWorkingName.ReadOnly = true;
			// 
			// colConditionDisplay
			// 
			this.colConditionDisplay.DataPropertyName = "Condition";
			this.colConditionDisplay.HeaderText = "Tiêu chuẩn";
			this.colConditionDisplay.Name = "colConditionDisplay";
			this.colConditionDisplay.ReadOnly = true;
			// 
			// colPeriodValue
			// 
			this.colPeriodValue.DataPropertyName = "PeriodValue";
			this.colPeriodValue.FillWeight = 76.4239F;
			this.colPeriodValue.HeaderText = "Tiêu chuẩn";
			this.colPeriodValue.Name = "colPeriodValue";
			this.colPeriodValue.ReadOnly = true;
			this.colPeriodValue.Visible = false;
			// 
			// colUnit
			// 
			this.colUnit.DataPropertyName = "Unit";
			this.colUnit.FillWeight = 68.41773F;
			this.colUnit.HeaderText = "Đơn vị";
			this.colUnit.Name = "colUnit";
			this.colUnit.ReadOnly = true;
			// 
			// colRealvalue
			// 
			this.colRealvalue.DataPropertyName = "Realvalue";
			this.colRealvalue.FillWeight = 70.32562F;
			this.colRealvalue.HeaderText = "Thực tế";
			this.colRealvalue.Name = "colRealvalue";
			// 
			// colEvaluate
			// 
			this.colEvaluate.DataPropertyName = "Evaluate";
			this.colEvaluate.FillWeight = 42.13918F;
			this.colEvaluate.HeaderText = "Đánh giá";
			this.colEvaluate.Name = "colEvaluate";
			this.colEvaluate.ReadOnly = true;
			// 
			// colY
			// 
			this.colY.DataPropertyName = "Y";
			this.colY.FillWeight = 45.68528F;
			this.colY.HeaderText = "Y";
			this.colY.Name = "colY";
			this.colY.Visible = false;
			// 
			// colX
			// 
			this.colX.DataPropertyName = "X";
			this.colX.HeaderText = "X";
			this.colX.Name = "colX";
			this.colX.Visible = false;
			// 
			// colMinValue
			// 
			this.colMinValue.DataPropertyName = "MinValue";
			this.colMinValue.HeaderText = "MinValue";
			this.colMinValue.Name = "colMinValue";
			this.colMinValue.Visible = false;
			// 
			// colID
			// 
			this.colID.DataPropertyName = "ID";
			this.colID.HeaderText = "ID";
			this.colID.Name = "colID";
			this.colID.Visible = false;
			// 
			// colMaxValue
			// 
			this.colMaxValue.DataPropertyName = "MaxValue";
			this.colMaxValue.FillWeight = 5.034656F;
			this.colMaxValue.HeaderText = "MaxValue";
			this.colMaxValue.Name = "colMaxValue";
			this.colMaxValue.Visible = false;
			// 
			// colDisplay
			// 
			this.colDisplay.DataPropertyName = "isDisplay";
			this.colDisplay.FillWeight = 34.20914F;
			this.colDisplay.HeaderText = "Hiển thị";
			this.colDisplay.Name = "colDisplay";
			this.colDisplay.Resizable = System.Windows.Forms.DataGridViewTriState.True;
			this.colDisplay.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
			this.colDisplay.Visible = false;
			// 
			// colCheckValueType
			// 
			this.colCheckValueType.DataPropertyName = "CheckValueType";
			this.colCheckValueType.HeaderText = "CheckValueType";
			this.colCheckValueType.Name = "colCheckValueType";
			this.colCheckValueType.Visible = false;
			// 
			// panel3
			// 
			this.panel3.Controls.Add(this.textBox1);
			this.panel3.Controls.Add(this.mnuMenu);
			this.panel3.Controls.Add(this.txtDateTime);
			this.panel3.Controls.Add(this.txtMachine);
			this.panel3.Controls.Add(this.cboCongDoan);
			this.panel3.Controls.Add(this.txtOrder);
			this.panel3.Controls.Add(this.txtMaHang);
			this.panel3.Controls.Add(this.txtWorkerCode);
			this.panel3.Controls.Add(this.button4);
			this.panel3.Controls.Add(this.button7);
			this.panel3.Controls.Add(this.button6);
			this.panel3.Controls.Add(this.button3);
			this.panel3.Controls.Add(this.button5);
			this.panel3.Controls.Add(this.button2);
			this.panel3.Controls.Add(this.button1);
			this.panel3.Dock = System.Windows.Forms.DockStyle.Fill;
			this.panel3.Location = new System.Drawing.Point(3, 3);
			this.panel3.Name = "panel3";
			this.panel3.Size = new System.Drawing.Size(1268, 58);
			this.panel3.TabIndex = 0;
			// 
			// textBox1
			// 
			this.textBox1.BackColor = System.Drawing.SystemColors.Window;
			this.textBox1.Font = new System.Drawing.Font("Calibri", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.textBox1.Location = new System.Drawing.Point(431, 2);
			this.textBox1.Name = "textBox1";
			this.textBox1.Size = new System.Drawing.Size(150, 26);
			this.textBox1.TabIndex = 74;
			// 
			// mnuMenu
			// 
			this.mnuMenu.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.mnuMenu.AutoSize = false;
			this.mnuMenu.BackColor = System.Drawing.SystemColors.ButtonHighlight;
			this.mnuMenu.Dock = System.Windows.Forms.DockStyle.None;
			this.mnuMenu.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.mnuMenu.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden;
			this.mnuMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.btnSave,
            this.btnPath});
			this.mnuMenu.LayoutStyle = System.Windows.Forms.ToolStripLayoutStyle.Flow;
			this.mnuMenu.Location = new System.Drawing.Point(869, 29);
			this.mnuMenu.Name = "mnuMenu";
			this.mnuMenu.RenderMode = System.Windows.Forms.ToolStripRenderMode.Professional;
			this.mnuMenu.Size = new System.Drawing.Size(396, 25);
			this.mnuMenu.TabIndex = 73;
			this.mnuMenu.Text = "toolStrip2";
			// 
			// btnSave
			// 
			this.btnSave.AutoSize = false;
			this.btnSave.Font = new System.Drawing.Font("Tahoma", 6.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnSave.Image = ((System.Drawing.Image)(resources.GetObject("btnSave.Image")));
			this.btnSave.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnSave.Name = "btnSave";
			this.btnSave.Size = new System.Drawing.Size(40, 20);
			this.btnSave.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
			this.btnSave.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
			// 
			// btnPath
			// 
			this.btnPath.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
			this.btnPath.Image = ((System.Drawing.Image)(resources.GetObject("btnPath.Image")));
			this.btnPath.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnPath.Name = "btnPath";
			this.btnPath.Size = new System.Drawing.Size(23, 20);
			this.btnPath.Text = "toolStripButton2";
			this.btnPath.Click += new System.EventHandler(this.btnPath_Click);
			// 
			// txtDateTime
			// 
			this.txtDateTime.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.txtDateTime.Font = new System.Drawing.Font("Calibri", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtDateTime.Location = new System.Drawing.Point(980, 1);
			this.txtDateTime.Name = "txtDateTime";
			this.txtDateTime.Size = new System.Drawing.Size(285, 26);
			this.txtDateTime.TabIndex = 4;
			// 
			// txtMachine
			// 
			this.txtMachine.Font = new System.Drawing.Font("Calibri", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtMachine.Location = new System.Drawing.Point(693, 29);
			this.txtMachine.Name = "txtMachine";
			this.txtMachine.Size = new System.Drawing.Size(175, 26);
			this.txtMachine.TabIndex = 5;
			this.txtMachine.TextChanged += new System.EventHandler(this.Textbox_TextChanged);
			// 
			// cboCongDoan
			// 
			this.cboCongDoan.Font = new System.Drawing.Font("Calibri", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.cboCongDoan.FormattingEnabled = true;
			this.cboCongDoan.Location = new System.Drawing.Point(693, 2);
			this.cboCongDoan.Name = "cboCongDoan";
			this.cboCongDoan.Size = new System.Drawing.Size(175, 26);
			this.cboCongDoan.TabIndex = 4;
			this.cboCongDoan.Text = "THAY KHUÔN";
			// 
			// txtOrder
			// 
			this.txtOrder.Font = new System.Drawing.Font("Calibri", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtOrder.Location = new System.Drawing.Point(114, 29);
			this.txtOrder.Name = "txtOrder";
			this.txtOrder.Size = new System.Drawing.Size(205, 26);
			this.txtOrder.TabIndex = 3;
			this.txtOrder.TextChanged += new System.EventHandler(this.Textbox_TextChanged);
			// 
			// txtMaHang
			// 
			this.txtMaHang.BackColor = System.Drawing.SystemColors.Window;
			this.txtMaHang.Font = new System.Drawing.Font("Calibri", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtMaHang.Location = new System.Drawing.Point(431, 29);
			this.txtMaHang.Name = "txtMaHang";
			this.txtMaHang.Size = new System.Drawing.Size(150, 26);
			this.txtMaHang.TabIndex = 2;
			this.txtMaHang.TextChanged += new System.EventHandler(this.Textbox_TextChanged);
			this.txtMaHang.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtMaHang_KeyDown);
			// 
			// txtWorkerCode
			// 
			this.txtWorkerCode.BackColor = System.Drawing.SystemColors.ButtonHighlight;
			this.txtWorkerCode.Font = new System.Drawing.Font("Calibri", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtWorkerCode.Location = new System.Drawing.Point(114, 2);
			this.txtWorkerCode.Name = "txtWorkerCode";
			this.txtWorkerCode.Size = new System.Drawing.Size(205, 26);
			this.txtWorkerCode.TabIndex = 2;
			// 
			// button4
			// 
			this.button4.BackColor = System.Drawing.Color.SkyBlue;
			this.button4.FlatAppearance.BorderColor = System.Drawing.Color.SkyBlue;
			this.button4.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
			this.button4.Font = new System.Drawing.Font("Calibri", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.button4.Location = new System.Drawing.Point(869, 2);
			this.button4.Name = "button4";
			this.button4.Size = new System.Drawing.Size(110, 25);
			this.button4.TabIndex = 0;
			this.button4.Text = "Thời gian";
			this.button4.UseVisualStyleBackColor = false;
			// 
			// button7
			// 
			this.button7.BackColor = System.Drawing.Color.SkyBlue;
			this.button7.FlatAppearance.BorderColor = System.Drawing.Color.SkyBlue;
			this.button7.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
			this.button7.Font = new System.Drawing.Font("Calibri", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.button7.Location = new System.Drawing.Point(582, 29);
			this.button7.Name = "button7";
			this.button7.Size = new System.Drawing.Size(110, 25);
			this.button7.TabIndex = 0;
			this.button7.Text = "Máy";
			this.button7.UseVisualStyleBackColor = false;
			// 
			// button6
			// 
			this.button6.BackColor = System.Drawing.Color.SkyBlue;
			this.button6.FlatAppearance.BorderColor = System.Drawing.Color.SkyBlue;
			this.button6.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
			this.button6.Font = new System.Drawing.Font("Calibri", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.button6.Location = new System.Drawing.Point(320, 29);
			this.button6.Name = "button6";
			this.button6.Size = new System.Drawing.Size(110, 25);
			this.button6.TabIndex = 0;
			this.button6.Text = "Mã hàng";
			this.button6.UseVisualStyleBackColor = false;
			// 
			// button3
			// 
			this.button3.BackColor = System.Drawing.Color.SkyBlue;
			this.button3.FlatAppearance.BorderColor = System.Drawing.Color.SkyBlue;
			this.button3.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
			this.button3.Font = new System.Drawing.Font("Calibri", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.button3.Location = new System.Drawing.Point(582, 3);
			this.button3.Name = "button3";
			this.button3.Size = new System.Drawing.Size(110, 25);
			this.button3.TabIndex = 0;
			this.button3.Text = "Công đoạn";
			this.button3.UseVisualStyleBackColor = false;
			// 
			// button5
			// 
			this.button5.BackColor = System.Drawing.Color.SkyBlue;
			this.button5.FlatAppearance.BorderColor = System.Drawing.Color.SkyBlue;
			this.button5.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
			this.button5.Font = new System.Drawing.Font("Calibri", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.button5.Location = new System.Drawing.Point(3, 29);
			this.button5.Name = "button5";
			this.button5.Size = new System.Drawing.Size(110, 25);
			this.button5.TabIndex = 0;
			this.button5.Text = "Số Order";
			this.button5.UseVisualStyleBackColor = false;
			// 
			// button2
			// 
			this.button2.BackColor = System.Drawing.Color.SkyBlue;
			this.button2.FlatAppearance.BorderColor = System.Drawing.Color.SkyBlue;
			this.button2.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
			this.button2.Font = new System.Drawing.Font("Calibri", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.button2.Location = new System.Drawing.Point(320, 3);
			this.button2.Name = "button2";
			this.button2.Size = new System.Drawing.Size(110, 25);
			this.button2.TabIndex = 0;
			this.button2.Text = "Bộ phận";
			this.button2.UseVisualStyleBackColor = false;
			// 
			// button1
			// 
			this.button1.BackColor = System.Drawing.Color.SkyBlue;
			this.button1.FlatAppearance.BorderColor = System.Drawing.Color.SkyBlue;
			this.button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
			this.button1.Font = new System.Drawing.Font("Calibri", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.button1.Location = new System.Drawing.Point(3, 3);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(110, 25);
			this.button1.TabIndex = 0;
			this.button1.Text = "Phụ trách";
			this.button1.UseVisualStyleBackColor = false;
			// 
			// frmHypMoldLine
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.BackColor = System.Drawing.Color.White;
			this.ClientSize = new System.Drawing.Size(1280, 961);
			this.Controls.Add(this.tableLayoutPanel1);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.MaximumSize = new System.Drawing.Size(1296, 1000);
			this.MinimumSize = new System.Drawing.Size(1278, 982);
			this.Name = "frmHypMoldLine";
			this.Text = "HYPOID PINION";
			this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
			this.Load += new System.EventHandler(this.Form1_Load);
			this.tableLayoutPanel1.ResumeLayout(false);
			this.panel2.ResumeLayout(false);
			this.panel2.PerformLayout();
			this.panel1.ResumeLayout(false);
			this.tableLayoutPanel2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.grvKnifeJigWorking)).EndInit();
			this.panel3.ResumeLayout(false);
			this.panel3.PerformLayout();
			this.mnuMenu.ResumeLayout(false);
			this.mnuMenu.PerformLayout();
			this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
        private DevExpress.XtraPdfViewer.PdfViewer pdfViewer1;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel2;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.TextBox txtDateTime;
        private System.Windows.Forms.TextBox txtMachine;
        private System.Windows.Forms.ComboBox cboCongDoan;
        private System.Windows.Forms.TextBox txtOrder;
        private System.Windows.Forms.TextBox txtMaHang;
        private System.Windows.Forms.TextBox txtWorkerCode;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Button button7;
        private System.Windows.Forms.Button button6;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.ToolStrip mnuMenu;
        private System.Windows.Forms.ToolStripButton btnSave;
        private System.Windows.Forms.DataGridView grvKnifeJigWorking;
        private System.Windows.Forms.Label lbly;
        private System.Windows.Forms.Label lblx;
        private System.Windows.Forms.ToolStripButton btnPath;
        private System.Windows.Forms.DataGridViewTextBoxColumn colSTT;
        private System.Windows.Forms.DataGridViewTextBoxColumn colWorkingName;
        private System.Windows.Forms.DataGridViewTextBoxColumn colConditionDisplay;
        private System.Windows.Forms.DataGridViewTextBoxColumn colPeriodValue;
        private System.Windows.Forms.DataGridViewTextBoxColumn colUnit;
        private System.Windows.Forms.DataGridViewTextBoxColumn colRealvalue;
        private System.Windows.Forms.DataGridViewTextBoxColumn colEvaluate;
        private System.Windows.Forms.DataGridViewTextBoxColumn colY;
        private System.Windows.Forms.DataGridViewTextBoxColumn colX;
        private System.Windows.Forms.DataGridViewTextBoxColumn colMinValue;
        private System.Windows.Forms.DataGridViewTextBoxColumn colID;
        private System.Windows.Forms.DataGridViewTextBoxColumn colMaxValue;
        private System.Windows.Forms.DataGridViewCheckBoxColumn colDisplay;
        private System.Windows.Forms.DataGridViewTextBoxColumn colCheckValueType;
		private System.Windows.Forms.TextBox textBox1;
	}
}

