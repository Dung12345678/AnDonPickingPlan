
namespace BMS
{
    partial class frmHypMold
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
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmHypMold));
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
			this.panel1 = new System.Windows.Forms.Panel();
			this.mnuMenu = new System.Windows.Forms.ToolStrip();
			this.btnZoom = new System.Windows.Forms.ToolStripButton();
			this.btnNew = new System.Windows.Forms.ToolStripButton();
			this.btnEdit = new System.Windows.Forms.ToolStripButton();
			this.btnDelete = new System.Windows.Forms.ToolStripButton();
			this.btnEditTxt = new System.Windows.Forms.ToolStripButton();
			this.btnPath = new System.Windows.Forms.ToolStripButton();
			this.btnSave = new System.Windows.Forms.ToolStripButton();
			this.btnCancel = new System.Windows.Forms.ToolStripButton();
			this.toolStripButton1 = new System.Windows.Forms.ToolStripButton();
			this.cboCongDoan = new System.Windows.Forms.ComboBox();
			this.txtMaHang = new System.Windows.Forms.TextBox();
			this.txtCreateby = new System.Windows.Forms.TextBox();
			this.button1 = new System.Windows.Forms.Button();
			this.button3 = new System.Windows.Forms.Button();
			this.button6 = new System.Windows.Forms.Button();
			this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
			this.panel2 = new System.Windows.Forms.Panel();
			this.lbly = new System.Windows.Forms.Label();
			this.lblx = new System.Windows.Forms.Label();
			this.pdfViewer1 = new DevExpress.XtraPdfViewer.PdfViewer();
			this.grvKnifeJigWorking = new System.Windows.Forms.DataGridView();
			this.colSTT = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.colWorkingName = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.colCondition = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.colPeriodValue = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.colUnit = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.colRealvalue = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.colEvaluate = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.colDisplay = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.colY = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.colX = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.colMinValue = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.colID = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.colMaxValue = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.panel1.SuspendLayout();
			this.mnuMenu.SuspendLayout();
			this.tableLayoutPanel1.SuspendLayout();
			this.panel2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.grvKnifeJigWorking)).BeginInit();
			this.SuspendLayout();
			// 
			// panel1
			// 
			this.panel1.Controls.Add(this.mnuMenu);
			this.panel1.Controls.Add(this.cboCongDoan);
			this.panel1.Controls.Add(this.txtMaHang);
			this.panel1.Controls.Add(this.txtCreateby);
			this.panel1.Controls.Add(this.button1);
			this.panel1.Controls.Add(this.button3);
			this.panel1.Controls.Add(this.button6);
			this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
			this.panel1.Location = new System.Drawing.Point(0, 0);
			this.panel1.Name = "panel1";
			this.panel1.Size = new System.Drawing.Size(1280, 74);
			this.panel1.TabIndex = 0;
			// 
			// mnuMenu
			// 
			this.mnuMenu.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.mnuMenu.AutoSize = false;
			this.mnuMenu.BackColor = System.Drawing.SystemColors.ButtonHighlight;
			this.mnuMenu.Dock = System.Windows.Forms.DockStyle.None;
			this.mnuMenu.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.mnuMenu.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden;
			this.mnuMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.btnZoom,
            this.btnNew,
            this.btnEdit,
            this.btnDelete,
            this.btnEditTxt,
            this.btnPath,
            this.btnSave,
            this.btnCancel,
            this.toolStripButton1});
			this.mnuMenu.LayoutStyle = System.Windows.Forms.ToolStripLayoutStyle.Flow;
			this.mnuMenu.Location = new System.Drawing.Point(655, 10);
			this.mnuMenu.Name = "mnuMenu";
			this.mnuMenu.RenderMode = System.Windows.Forms.ToolStripRenderMode.Professional;
			this.mnuMenu.Size = new System.Drawing.Size(409, 44);
			this.mnuMenu.TabIndex = 73;
			this.mnuMenu.Text = "toolStrip2";
			// 
			// btnZoom
			// 
			this.btnZoom.AutoSize = false;
			this.btnZoom.Font = new System.Drawing.Font("Tahoma", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnZoom.Image = ((System.Drawing.Image)(resources.GetObject("btnZoom.Image")));
			this.btnZoom.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnZoom.Name = "btnZoom";
			this.btnZoom.Size = new System.Drawing.Size(20, 20);
			this.btnZoom.TextAlign = System.Drawing.ContentAlignment.TopCenter;
			this.btnZoom.Click += new System.EventHandler(this.btnZoom_Click);
			// 
			// btnNew
			// 
			this.btnNew.AutoSize = false;
			this.btnNew.Font = new System.Drawing.Font("Tahoma", 6.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnNew.Image = ((System.Drawing.Image)(resources.GetObject("btnNew.Image")));
			this.btnNew.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnNew.Name = "btnNew";
			this.btnNew.Size = new System.Drawing.Size(40, 20);
			this.btnNew.Tag = "";
			this.btnNew.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnNew.Click += new System.EventHandler(this.btnNew_Click);
			// 
			// btnEdit
			// 
			this.btnEdit.AutoSize = false;
			this.btnEdit.Font = new System.Drawing.Font("Tahoma", 6.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnEdit.Image = ((System.Drawing.Image)(resources.GetObject("btnEdit.Image")));
			this.btnEdit.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnEdit.Name = "btnEdit";
			this.btnEdit.Size = new System.Drawing.Size(40, 20);
			this.btnEdit.Tag = "";
			this.btnEdit.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnEdit.Click += new System.EventHandler(this.btnEdit_Click);
			// 
			// btnDelete
			// 
			this.btnDelete.AutoSize = false;
			this.btnDelete.Font = new System.Drawing.Font("Tahoma", 6.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnDelete.Image = ((System.Drawing.Image)(resources.GetObject("btnDelete.Image")));
			this.btnDelete.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnDelete.Name = "btnDelete";
			this.btnDelete.Size = new System.Drawing.Size(40, 20);
			this.btnDelete.Tag = "";
			this.btnDelete.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnDelete.Click += new System.EventHandler(this.btnDelete_Click);
			// 
			// btnEditTxt
			// 
			this.btnEditTxt.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
			this.btnEditTxt.Image = ((System.Drawing.Image)(resources.GetObject("btnEditTxt.Image")));
			this.btnEditTxt.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnEditTxt.Name = "btnEditTxt";
			this.btnEditTxt.Size = new System.Drawing.Size(23, 20);
			this.btnEditTxt.Text = "toolStripButton1";
			this.btnEditTxt.Click += new System.EventHandler(this.btnEditTxt_Click);
			// 
			// btnPath
			// 
			this.btnPath.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
			this.btnPath.Image = ((System.Drawing.Image)(resources.GetObject("btnPath.Image")));
			this.btnPath.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnPath.Name = "btnPath";
			this.btnPath.Size = new System.Drawing.Size(23, 20);
			this.btnPath.Text = "toolStripButton2";
			this.btnPath.Click += new System.EventHandler(this.toolStripButton2_Click);
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
			// btnCancel
			// 
			this.btnCancel.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
			this.btnCancel.AutoSize = false;
			this.btnCancel.Image = ((System.Drawing.Image)(resources.GetObject("btnCancel.Image")));
			this.btnCancel.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnCancel.Name = "btnCancel";
			this.btnCancel.Size = new System.Drawing.Size(40, 20);
			this.btnCancel.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
			// 
			// toolStripButton1
			// 
			this.toolStripButton1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
			this.toolStripButton1.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.toolStripButton1.Name = "toolStripButton1";
			this.toolStripButton1.Size = new System.Drawing.Size(23, 4);
			this.toolStripButton1.Text = "toolStripButton1";
			// 
			// cboCongDoan
			// 
			this.cboCongDoan.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.cboCongDoan.Font = new System.Drawing.Font("Calibri", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.cboCongDoan.FormattingEnabled = true;
			this.cboCongDoan.Location = new System.Drawing.Point(462, 9);
			this.cboCongDoan.Name = "cboCongDoan";
			this.cboCongDoan.Size = new System.Drawing.Size(181, 26);
			this.cboCongDoan.TabIndex = 4;
			this.cboCongDoan.Text = "THAY KHUÔN";
			this.cboCongDoan.SelectedIndexChanged += new System.EventHandler(this.cboCongDoan_SelectedIndexChanged);
			// 
			// txtMaHang
			// 
			this.txtMaHang.Font = new System.Drawing.Font("Calibri", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtMaHang.Location = new System.Drawing.Point(129, 41);
			this.txtMaHang.Name = "txtMaHang";
			this.txtMaHang.Size = new System.Drawing.Size(205, 26);
			this.txtMaHang.TabIndex = 2;
			this.txtMaHang.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtMaHang_KeyDown);
			// 
			// txtCreateby
			// 
			this.txtCreateby.Font = new System.Drawing.Font("Calibri", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtCreateby.Location = new System.Drawing.Point(129, 9);
			this.txtCreateby.Name = "txtCreateby";
			this.txtCreateby.Size = new System.Drawing.Size(205, 26);
			this.txtCreateby.TabIndex = 2;
			this.txtCreateby.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCreateby_KeyDown);
			// 
			// button1
			// 
			this.button1.BackColor = System.Drawing.Color.SkyBlue;
			this.button1.FlatAppearance.BorderColor = System.Drawing.Color.SkyBlue;
			this.button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
			this.button1.Font = new System.Drawing.Font("Calibri", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.button1.Location = new System.Drawing.Point(13, 10);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(110, 25);
			this.button1.TabIndex = 0;
			this.button1.Text = "Phụ trách";
			this.button1.UseVisualStyleBackColor = true;
			// 
			// button3
			// 
			this.button3.BackColor = System.Drawing.Color.SkyBlue;
			this.button3.FlatAppearance.BorderColor = System.Drawing.Color.SkyBlue;
			this.button3.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
			this.button3.Font = new System.Drawing.Font("Calibri", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.button3.Location = new System.Drawing.Point(346, 10);
			this.button3.Name = "button3";
			this.button3.Size = new System.Drawing.Size(110, 25);
			this.button3.TabIndex = 0;
			this.button3.Text = "Công đoạn";
			this.button3.UseVisualStyleBackColor = true;
			// 
			// button6
			// 
			this.button6.BackColor = System.Drawing.Color.SkyBlue;
			this.button6.FlatAppearance.BorderColor = System.Drawing.Color.SkyBlue;
			this.button6.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
			this.button6.Font = new System.Drawing.Font("Calibri", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.button6.Location = new System.Drawing.Point(13, 42);
			this.button6.Name = "button6";
			this.button6.Size = new System.Drawing.Size(110, 25);
			this.button6.TabIndex = 0;
			this.button6.Text = "Mã hàng";
			this.button6.UseVisualStyleBackColor = true;
			// 
			// tableLayoutPanel1
			// 
			this.tableLayoutPanel1.ColumnCount = 1;
			this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 1000F));
			this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 20F));
			this.tableLayoutPanel1.Controls.Add(this.panel2, 0, 0);
			this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 71);
			this.tableLayoutPanel1.Name = "tableLayoutPanel1";
			this.tableLayoutPanel1.RowCount = 1;
			this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
			this.tableLayoutPanel1.Size = new System.Drawing.Size(1288, 839);
			this.tableLayoutPanel1.TabIndex = 0;
			// 
			// panel2
			// 
			this.panel2.Controls.Add(this.lbly);
			this.panel2.Controls.Add(this.lblx);
			this.panel2.Controls.Add(this.pdfViewer1);
			this.panel2.Controls.Add(this.grvKnifeJigWorking);
			this.panel2.Location = new System.Drawing.Point(3, 3);
			this.panel2.Name = "panel2";
			this.panel2.Size = new System.Drawing.Size(1282, 833);
			this.panel2.TabIndex = 1;
			// 
			// lbly
			// 
			this.lbly.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lbly.AutoSize = true;
			this.lbly.BackColor = System.Drawing.SystemColors.Control;
			this.lbly.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lbly.Location = new System.Drawing.Point(1224, 3);
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
			this.lblx.Location = new System.Drawing.Point(1184, 3);
			this.lblx.Name = "lblx";
			this.lblx.Size = new System.Drawing.Size(14, 16);
			this.lblx.TabIndex = 0;
			this.lblx.Text = "x";
			// 
			// pdfViewer1
			// 
			this.pdfViewer1.Cursor = System.Windows.Forms.Cursors.Default;
			this.pdfViewer1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.pdfViewer1.Location = new System.Drawing.Point(0, 0);
			this.pdfViewer1.LookAndFeel.SkinName = "VS2010";
			this.pdfViewer1.LookAndFeel.Style = DevExpress.LookAndFeel.LookAndFeelStyle.UltraFlat;
			this.pdfViewer1.MaximumSize = new System.Drawing.Size(1277, 834);
			this.pdfViewer1.MinimumSize = new System.Drawing.Size(1277, 834);
			this.pdfViewer1.Name = "pdfViewer1";
			this.pdfViewer1.Size = new System.Drawing.Size(1277, 834);
			this.pdfViewer1.TabIndex = 0;
			this.pdfViewer1.ZoomMode = DevExpress.XtraPdfViewer.PdfZoomMode.PageLevel;
			// 
			// grvKnifeJigWorking
			// 
			this.grvKnifeJigWorking.AllowUserToAddRows = false;
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
            this.colCondition,
            this.colPeriodValue,
            this.colUnit,
            this.colRealvalue,
            this.colEvaluate,
            this.colDisplay,
            this.colY,
            this.colX,
            this.colMinValue,
            this.colID,
            this.colMaxValue});
			dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
			dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window;
			dataGridViewCellStyle2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText;
			dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
			dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
			dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
			this.grvKnifeJigWorking.DefaultCellStyle = dataGridViewCellStyle2;
			this.grvKnifeJigWorking.EnableHeadersVisualStyles = false;
			this.grvKnifeJigWorking.Location = new System.Drawing.Point(-3, 0);
			this.grvKnifeJigWorking.Name = "grvKnifeJigWorking";
			this.grvKnifeJigWorking.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.Single;
			this.grvKnifeJigWorking.RowHeadersVisible = false;
			this.grvKnifeJigWorking.RowHeadersWidth = 10;
			this.grvKnifeJigWorking.RowTemplate.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
			this.grvKnifeJigWorking.RowTemplate.DefaultCellStyle.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.grvKnifeJigWorking.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
			this.grvKnifeJigWorking.Size = new System.Drawing.Size(1280, 833);
			this.grvKnifeJigWorking.TabIndex = 75;
			// 
			// colSTT
			// 
			this.colSTT.DataPropertyName = "SortOrder";
			this.colSTT.FillWeight = 25.57112F;
			this.colSTT.HeaderText = "STT";
			this.colSTT.Name = "colSTT";
			this.colSTT.Visible = false;
			// 
			// colWorkingName
			// 
			this.colWorkingName.DataPropertyName = "WorkingName";
			this.colWorkingName.FillWeight = 107.1869F;
			this.colWorkingName.HeaderText = "Mục kiểm tra";
			this.colWorkingName.Name = "colWorkingName";
			this.colWorkingName.Width = 598;
			// 
			// colCondition
			// 
			this.colCondition.DataPropertyName = "Condition";
			this.colCondition.HeaderText = "Tiêu chuẩn";
			this.colCondition.Name = "colCondition";
			this.colCondition.Visible = false;
			// 
			// colPeriodValue
			// 
			this.colPeriodValue.DataPropertyName = "PeriodValue";
			this.colPeriodValue.FillWeight = 64.31032F;
			this.colPeriodValue.HeaderText = "Tiêu chuẩn";
			this.colPeriodValue.Name = "colPeriodValue";
			this.colPeriodValue.Width = 359;
			// 
			// colUnit
			// 
			this.colUnit.DataPropertyName = "Unit";
			this.colUnit.FillWeight = 57.57317F;
			this.colUnit.HeaderText = "Đơn vị";
			this.colUnit.Name = "colUnit";
			this.colUnit.Width = 320;
			// 
			// colRealvalue
			// 
			this.colRealvalue.DataPropertyName = "Realvalue";
			this.colRealvalue.FillWeight = 59.17865F;
			this.colRealvalue.HeaderText = "Thực tế";
			this.colRealvalue.Name = "colRealvalue";
			this.colRealvalue.Visible = false;
			// 
			// colEvaluate
			// 
			this.colEvaluate.DataPropertyName = "Evaluate";
			this.colEvaluate.FillWeight = 35.4599F;
			this.colEvaluate.HeaderText = "Đánh giá";
			this.colEvaluate.Name = "colEvaluate";
			this.colEvaluate.Visible = false;
			// 
			// colDisplay
			// 
			this.colDisplay.DataPropertyName = "isDisplay";
			this.colDisplay.HeaderText = "Display";
			this.colDisplay.Name = "colDisplay";
			this.colDisplay.Visible = false;
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
			// frmHypMold
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.BackColor = System.Drawing.Color.White;
			this.ClientSize = new System.Drawing.Size(1280, 943);
			this.Controls.Add(this.panel1);
			this.Controls.Add(this.tableLayoutPanel1);
			this.MaximumSize = new System.Drawing.Size(1296, 1000);
			this.MinimumSize = new System.Drawing.Size(1278, 982);
			this.Name = "frmHypMold";
			this.Text = "HYPOID PINION";
			this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
			this.Load += new System.EventHandler(this.Form1_Load);
			this.panel1.ResumeLayout(false);
			this.panel1.PerformLayout();
			this.mnuMenu.ResumeLayout(false);
			this.mnuMenu.PerformLayout();
			this.tableLayoutPanel1.ResumeLayout(false);
			this.panel2.ResumeLayout(false);
			this.panel2.PerformLayout();
			((System.ComponentModel.ISupportInitialize)(this.grvKnifeJigWorking)).EndInit();
			this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.ComboBox cboCongDoan;
        private System.Windows.Forms.TextBox txtMaHang;
        private System.Windows.Forms.TextBox txtCreateby;
        private System.Windows.Forms.Button button6;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.ToolStrip mnuMenu;
        private System.Windows.Forms.ToolStripButton btnNew;
        private System.Windows.Forms.ToolStripButton btnEdit;
        private System.Windows.Forms.ToolStripButton btnDelete;
        private System.Windows.Forms.ToolStripButton btnSave;
        private System.Windows.Forms.ToolStripButton btnCancel;
        private System.Windows.Forms.ToolStripButton btnEditTxt;
        private System.Windows.Forms.ToolStripButton toolStripButton1;
        private System.Windows.Forms.ToolStripButton btnPath;
		private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
		private System.Windows.Forms.DataGridView grvKnifeJigWorking;
		private System.Windows.Forms.DataGridViewTextBoxColumn colSTT;
		private System.Windows.Forms.DataGridViewTextBoxColumn colWorkingName;
		private System.Windows.Forms.DataGridViewTextBoxColumn colCondition;
		private System.Windows.Forms.DataGridViewTextBoxColumn colPeriodValue;
		private System.Windows.Forms.DataGridViewTextBoxColumn colUnit;
		private System.Windows.Forms.DataGridViewTextBoxColumn colRealvalue;
		private System.Windows.Forms.DataGridViewTextBoxColumn colEvaluate;
		private System.Windows.Forms.DataGridViewTextBoxColumn colDisplay;
		private System.Windows.Forms.DataGridViewTextBoxColumn colY;
		private System.Windows.Forms.DataGridViewTextBoxColumn colX;
		private System.Windows.Forms.DataGridViewTextBoxColumn colMinValue;
		private System.Windows.Forms.DataGridViewTextBoxColumn colID;
		private System.Windows.Forms.DataGridViewTextBoxColumn colMaxValue;
		private System.Windows.Forms.Panel panel2;
		private System.Windows.Forms.Label lbly;
		private System.Windows.Forms.Label lblx;
		private DevExpress.XtraPdfViewer.PdfViewer pdfViewer1;
		private System.Windows.Forms.ToolStripButton btnZoom;
	}
}

