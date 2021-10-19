
namespace BMS
{
    partial class frmHypMoldLineInformation
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
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmHypMoldLineInformation));
			this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
			this.panel2 = new System.Windows.Forms.Panel();
			this.lbly = new System.Windows.Forms.Label();
			this.lblx = new System.Windows.Forms.Label();
			this.pdfViewer1 = new DevExpress.XtraPdfViewer.PdfViewer();
			this.panel3 = new System.Windows.Forms.Panel();
			this.grdData = new DevExpress.XtraGrid.GridControl();
			this.grvData = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.colWorkingName = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colPeriodValue = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colRealvalue = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colSortOrder = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colCondition = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colUnit = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colResult = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colX = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colY = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colMinValue = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colMaxValue = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colisDisplay = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colCheckValueType = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colID = new DevExpress.XtraGrid.Columns.GridColumn();
			this.panel1 = new System.Windows.Forms.Panel();
			this.label2 = new System.Windows.Forms.Label();
			this.cboPart = new System.Windows.Forms.ComboBox();
			this.txtGoodsCode = new System.Windows.Forms.TextBox();
			this.label7 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.mnuMenu = new System.Windows.Forms.ToolStrip();
			this.btnSave = new System.Windows.Forms.ToolStripButton();
			this.btnPath = new System.Windows.Forms.ToolStripButton();
			this.txtMachine = new System.Windows.Forms.TextBox();
			this.cboCongDoan = new System.Windows.Forms.ComboBox();
			this.txtOrder = new System.Windows.Forms.TextBox();
			this.txtKnifeDetailCode = new System.Windows.Forms.TextBox();
			this.txtWorkerCode = new System.Windows.Forms.TextBox();
			this.tableLayoutPanel1.SuspendLayout();
			this.panel2.SuspendLayout();
			this.panel3.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.grdData)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grvData)).BeginInit();
			this.panel1.SuspendLayout();
			this.mnuMenu.SuspendLayout();
			this.SuspendLayout();
			// 
			// tableLayoutPanel1
			// 
			this.tableLayoutPanel1.ColumnCount = 2;
			this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
			this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 1000F));
			this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 20F));
			this.tableLayoutPanel1.Controls.Add(this.panel2, 1, 0);
			this.tableLayoutPanel1.Controls.Add(this.panel3, 0, 0);
			this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 83);
			this.tableLayoutPanel1.MaximumSize = new System.Drawing.Size(1280, 878);
			this.tableLayoutPanel1.MinimumSize = new System.Drawing.Size(1280, 878);
			this.tableLayoutPanel1.Name = "tableLayoutPanel1";
			this.tableLayoutPanel1.RowCount = 1;
			this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 23.62123F));
			this.tableLayoutPanel1.Size = new System.Drawing.Size(1280, 878);
			this.tableLayoutPanel1.TabIndex = 0;
			// 
			// panel2
			// 
			this.panel2.Controls.Add(this.lbly);
			this.panel2.Controls.Add(this.lblx);
			this.panel2.Controls.Add(this.pdfViewer1);
			this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
			this.panel2.Location = new System.Drawing.Point(283, 3);
			this.panel2.Name = "panel2";
			this.panel2.Size = new System.Drawing.Size(994, 872);
			this.panel2.TabIndex = 1;
			// 
			// lbly
			// 
			this.lbly.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lbly.AutoSize = true;
			this.lbly.BackColor = System.Drawing.SystemColors.Control;
			this.lbly.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lbly.Location = new System.Drawing.Point(937, 12);
			this.lbly.Name = "lbly";
			this.lbly.Size = new System.Drawing.Size(15, 16);
			this.lbly.TabIndex = 77;
			this.lbly.Text = "y";
			// 
			// lblx
			// 
			this.lblx.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblx.AutoSize = true;
			this.lblx.BackColor = System.Drawing.SystemColors.Control;
			this.lblx.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblx.Location = new System.Drawing.Point(888, 12);
			this.lblx.Name = "lblx";
			this.lblx.Size = new System.Drawing.Size(14, 16);
			this.lblx.TabIndex = 75;
			this.lblx.Text = "x";
			// 
			// pdfViewer1
			// 
			this.pdfViewer1.Cursor = System.Windows.Forms.Cursors.Default;
			this.pdfViewer1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.pdfViewer1.Location = new System.Drawing.Point(0, 0);
			this.pdfViewer1.LookAndFeel.SkinName = "VS2010";
			this.pdfViewer1.LookAndFeel.Style = DevExpress.LookAndFeel.LookAndFeelStyle.UltraFlat;
			this.pdfViewer1.MaximumSize = new System.Drawing.Size(994, 872);
			this.pdfViewer1.MinimumSize = new System.Drawing.Size(994, 872);
			this.pdfViewer1.Name = "pdfViewer1";
			this.pdfViewer1.Size = new System.Drawing.Size(994, 872);
			this.pdfViewer1.TabIndex = 76;
			this.pdfViewer1.ZoomMode = DevExpress.XtraPdfViewer.PdfZoomMode.PageLevel;
			// 
			// panel3
			// 
			this.panel3.Controls.Add(this.grdData);
			this.panel3.Dock = System.Windows.Forms.DockStyle.Fill;
			this.panel3.Location = new System.Drawing.Point(3, 3);
			this.panel3.Name = "panel3";
			this.panel3.Size = new System.Drawing.Size(274, 872);
			this.panel3.TabIndex = 0;
			// 
			// grdData
			// 
			this.grdData.Dock = System.Windows.Forms.DockStyle.Fill;
			this.grdData.Location = new System.Drawing.Point(0, 0);
			this.grdData.MainView = this.grvData;
			this.grdData.Name = "grdData";
			this.grdData.Size = new System.Drawing.Size(274, 872);
			this.grdData.TabIndex = 0;
			this.grdData.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.grvData});
			// 
			// grvData
			// 
			this.grvData.Appearance.HeaderPanel.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(51)))), ((int)(((byte)(204)))));
			this.grvData.Appearance.HeaderPanel.ForeColor = System.Drawing.Color.White;
			this.grvData.Appearance.HeaderPanel.Options.UseBackColor = true;
			this.grvData.Appearance.HeaderPanel.Options.UseForeColor = true;
			this.grvData.ColumnPanelRowHeight = 30;
			this.grvData.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
            this.colWorkingName,
            this.colPeriodValue,
            this.colRealvalue,
            this.colSortOrder,
            this.colCondition,
            this.colUnit,
            this.colResult,
            this.colX,
            this.colY,
            this.colMinValue,
            this.colMaxValue,
            this.colisDisplay,
            this.colCheckValueType,
            this.colID});
			this.grvData.GridControl = this.grdData;
			this.grvData.Name = "grvData";
			this.grvData.OptionsView.ShowDetailButtons = false;
			this.grvData.OptionsView.ShowGroupExpandCollapseButtons = false;
			this.grvData.OptionsView.ShowGroupPanel = false;
			this.grvData.PaintStyleName = "Flat";
			this.grvData.RowCellStyle += new DevExpress.XtraGrid.Views.Grid.RowCellStyleEventHandler(this.grvData_RowCellStyle);
			// 
			// colWorkingName
			// 
			this.colWorkingName.AppearanceCell.Options.UseTextOptions = true;
			this.colWorkingName.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colWorkingName.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colWorkingName.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.colWorkingName.AppearanceHeader.Options.UseFont = true;
			this.colWorkingName.AppearanceHeader.Options.UseTextOptions = true;
			this.colWorkingName.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colWorkingName.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colWorkingName.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colWorkingName.Caption = "Mục kiểm tra";
			this.colWorkingName.FieldName = "WorkingName";
			this.colWorkingName.Name = "colWorkingName";
			this.colWorkingName.OptionsColumn.AllowEdit = false;
			this.colWorkingName.OptionsColumn.AllowFocus = false;
			this.colWorkingName.Visible = true;
			this.colWorkingName.VisibleIndex = 0;
			// 
			// colPeriodValue
			// 
			this.colPeriodValue.AppearanceCell.Options.UseTextOptions = true;
			this.colPeriodValue.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colPeriodValue.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colPeriodValue.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.colPeriodValue.AppearanceHeader.Options.UseFont = true;
			this.colPeriodValue.AppearanceHeader.Options.UseTextOptions = true;
			this.colPeriodValue.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colPeriodValue.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colPeriodValue.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colPeriodValue.Caption = "Tiêu chuẩn ";
			this.colPeriodValue.FieldName = "StandardValue";
			this.colPeriodValue.Name = "colPeriodValue";
			this.colPeriodValue.OptionsColumn.AllowEdit = false;
			this.colPeriodValue.OptionsColumn.AllowFocus = false;
			this.colPeriodValue.Visible = true;
			this.colPeriodValue.VisibleIndex = 1;
			// 
			// colRealvalue
			// 
			this.colRealvalue.AppearanceCell.Options.UseTextOptions = true;
			this.colRealvalue.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colRealvalue.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colRealvalue.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colRealvalue.AppearanceHeader.Font = new System.Drawing.Font("Calibri", 11.25F, System.Drawing.FontStyle.Bold);
			this.colRealvalue.AppearanceHeader.Options.UseFont = true;
			this.colRealvalue.AppearanceHeader.Options.UseTextOptions = true;
			this.colRealvalue.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colRealvalue.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colRealvalue.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colRealvalue.Caption = "Thực tế";
			this.colRealvalue.FieldName = "RealValue";
			this.colRealvalue.Name = "colRealvalue";
			this.colRealvalue.Visible = true;
			this.colRealvalue.VisibleIndex = 2;
			// 
			// colSortOrder
			// 
			this.colSortOrder.Caption = "STT";
			this.colSortOrder.FieldName = "SortOrder";
			this.colSortOrder.Name = "colSortOrder";
			// 
			// colCondition
			// 
			this.colCondition.Caption = "Tiêu chuẩn";
			this.colCondition.FieldName = "Condition";
			this.colCondition.Name = "colCondition";
			// 
			// colUnit
			// 
			this.colUnit.Caption = "Đơn vị";
			this.colUnit.FieldName = "Unit";
			this.colUnit.Name = "colUnit";
			// 
			// colResult
			// 
			this.colResult.Caption = "Đánh giá";
			this.colResult.FieldName = "Result";
			this.colResult.Name = "colResult";
			// 
			// colX
			// 
			this.colX.Caption = "X";
			this.colX.FieldName = "X";
			this.colX.Name = "colX";
			// 
			// colY
			// 
			this.colY.Caption = "Y";
			this.colY.FieldName = "Y";
			this.colY.Name = "colY";
			// 
			// colMinValue
			// 
			this.colMinValue.Caption = "Min Value";
			this.colMinValue.FieldName = "MinValue";
			this.colMinValue.Name = "colMinValue";
			// 
			// colMaxValue
			// 
			this.colMaxValue.Caption = "Max Value";
			this.colMaxValue.FieldName = "MaxValue";
			this.colMaxValue.Name = "colMaxValue";
			// 
			// colisDisplay
			// 
			this.colisDisplay.Caption = "Hiển thị";
			this.colisDisplay.FieldName = "isDisplay";
			this.colisDisplay.Name = "colisDisplay";
			// 
			// colCheckValueType
			// 
			this.colCheckValueType.Caption = "Check Value Type";
			this.colCheckValueType.FieldName = "CheckValueType";
			this.colCheckValueType.Name = "colCheckValueType";
			// 
			// colID
			// 
			this.colID.Caption = "ID";
			this.colID.FieldName = "ID";
			this.colID.Name = "colID";
			// 
			// panel1
			// 
			this.panel1.Controls.Add(this.label2);
			this.panel1.Controls.Add(this.cboPart);
			this.panel1.Controls.Add(this.txtGoodsCode);
			this.panel1.Controls.Add(this.label7);
			this.panel1.Controls.Add(this.label6);
			this.panel1.Controls.Add(this.label5);
			this.panel1.Controls.Add(this.label4);
			this.panel1.Controls.Add(this.label3);
			this.panel1.Controls.Add(this.label1);
			this.panel1.Controls.Add(this.mnuMenu);
			this.panel1.Controls.Add(this.txtMachine);
			this.panel1.Controls.Add(this.cboCongDoan);
			this.panel1.Controls.Add(this.txtOrder);
			this.panel1.Controls.Add(this.txtKnifeDetailCode);
			this.panel1.Controls.Add(this.txtWorkerCode);
			this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
			this.panel1.Location = new System.Drawing.Point(0, 0);
			this.panel1.Name = "panel1";
			this.panel1.Size = new System.Drawing.Size(1280, 83);
			this.panel1.TabIndex = 1;
			// 
			// label2
			// 
			this.label2.BackColor = System.Drawing.Color.SkyBlue;
			this.label2.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
			this.label2.Font = new System.Drawing.Font("Calibri", 11.25F, System.Drawing.FontStyle.Bold);
			this.label2.Location = new System.Drawing.Point(305, 44);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(110, 25);
			this.label2.TabIndex = 113;
			this.label2.Text = "Mã hàng";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// cboPart
			// 
			this.cboPart.Font = new System.Drawing.Font("Calibri", 11.25F, System.Drawing.FontStyle.Bold);
			this.cboPart.FormattingEnabled = true;
			this.cboPart.Location = new System.Drawing.Point(421, 12);
			this.cboPart.Name = "cboPart";
			this.cboPart.Size = new System.Drawing.Size(172, 26);
			this.cboPart.TabIndex = 114;
			// 
			// txtGoodsCode
			// 
			this.txtGoodsCode.BackColor = System.Drawing.SystemColors.Window;
			this.txtGoodsCode.Font = new System.Drawing.Font("Calibri", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtGoodsCode.Location = new System.Drawing.Point(421, 44);
			this.txtGoodsCode.Name = "txtGoodsCode";
			this.txtGoodsCode.Size = new System.Drawing.Size(172, 26);
			this.txtGoodsCode.TabIndex = 112;
			// 
			// label7
			// 
			this.label7.BackColor = System.Drawing.Color.SkyBlue;
			this.label7.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
			this.label7.Font = new System.Drawing.Font("Calibri", 9.75F, System.Drawing.FontStyle.Bold);
			this.label7.Location = new System.Drawing.Point(11, 41);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(110, 25);
			this.label7.TabIndex = 111;
			this.label7.Text = "Số Order";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label6
			// 
			this.label6.BackColor = System.Drawing.Color.SkyBlue;
			this.label6.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
			this.label6.Font = new System.Drawing.Font("Calibri", 9.75F, System.Drawing.FontStyle.Bold);
			this.label6.Location = new System.Drawing.Point(305, 12);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(110, 25);
			this.label6.TabIndex = 110;
			this.label6.Text = "Bộ phận";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label5
			// 
			this.label5.BackColor = System.Drawing.Color.SkyBlue;
			this.label5.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
			this.label5.Font = new System.Drawing.Font("Calibri", 9.75F, System.Drawing.FontStyle.Bold);
			this.label5.Location = new System.Drawing.Point(599, 42);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(110, 25);
			this.label5.TabIndex = 109;
			this.label5.Text = "Mã khuôn";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label4
			// 
			this.label4.BackColor = System.Drawing.Color.SkyBlue;
			this.label4.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
			this.label4.Font = new System.Drawing.Font("Calibri", 9.75F, System.Drawing.FontStyle.Bold);
			this.label4.Location = new System.Drawing.Point(897, 42);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(110, 25);
			this.label4.TabIndex = 108;
			this.label4.Text = "Công đoạn";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label3
			// 
			this.label3.BackColor = System.Drawing.Color.SkyBlue;
			this.label3.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
			this.label3.Font = new System.Drawing.Font("Calibri", 9.75F, System.Drawing.FontStyle.Bold);
			this.label3.Location = new System.Drawing.Point(599, 12);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(110, 25);
			this.label3.TabIndex = 107;
			this.label3.Text = "Máy";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label1
			// 
			this.label1.BackColor = System.Drawing.Color.SkyBlue;
			this.label1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
			this.label1.Font = new System.Drawing.Font("Calibri", 9.75F, System.Drawing.FontStyle.Bold);
			this.label1.Location = new System.Drawing.Point(11, 12);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(110, 25);
			this.label1.TabIndex = 105;
			this.label1.Text = "Phụ trách";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
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
			this.mnuMenu.Location = new System.Drawing.Point(1159, 43);
			this.mnuMenu.Name = "mnuMenu";
			this.mnuMenu.RenderMode = System.Windows.Forms.ToolStripRenderMode.Professional;
			this.mnuMenu.Size = new System.Drawing.Size(112, 25);
			this.mnuMenu.TabIndex = 103;
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
			this.btnSave.Visible = false;
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
			// txtMachine
			// 
			this.txtMachine.Font = new System.Drawing.Font("Calibri", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtMachine.Location = new System.Drawing.Point(715, 11);
			this.txtMachine.Name = "txtMachine";
			this.txtMachine.Size = new System.Drawing.Size(172, 26);
			this.txtMachine.TabIndex = 102;
			// 
			// cboCongDoan
			// 
			this.cboCongDoan.Font = new System.Drawing.Font("Calibri", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.cboCongDoan.FormattingEnabled = true;
			this.cboCongDoan.Location = new System.Drawing.Point(1013, 43);
			this.cboCongDoan.Name = "cboCongDoan";
			this.cboCongDoan.Size = new System.Drawing.Size(143, 26);
			this.cboCongDoan.TabIndex = 101;
			// 
			// txtOrder
			// 
			this.txtOrder.Font = new System.Drawing.Font("Calibri", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtOrder.Location = new System.Drawing.Point(127, 41);
			this.txtOrder.Name = "txtOrder";
			this.txtOrder.Size = new System.Drawing.Size(172, 26);
			this.txtOrder.TabIndex = 99;
			// 
			// txtKnifeDetailCode
			// 
			this.txtKnifeDetailCode.BackColor = System.Drawing.SystemColors.Window;
			this.txtKnifeDetailCode.Font = new System.Drawing.Font("Calibri", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtKnifeDetailCode.Location = new System.Drawing.Point(715, 41);
			this.txtKnifeDetailCode.Name = "txtKnifeDetailCode";
			this.txtKnifeDetailCode.Size = new System.Drawing.Size(172, 26);
			this.txtKnifeDetailCode.TabIndex = 97;
			// 
			// txtWorkerCode
			// 
			this.txtWorkerCode.BackColor = System.Drawing.SystemColors.ButtonHighlight;
			this.txtWorkerCode.Font = new System.Drawing.Font("Calibri", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtWorkerCode.Location = new System.Drawing.Point(127, 12);
			this.txtWorkerCode.Name = "txtWorkerCode";
			this.txtWorkerCode.Size = new System.Drawing.Size(172, 26);
			this.txtWorkerCode.TabIndex = 98;
			// 
			// frmHypMoldLineInformation
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.BackColor = System.Drawing.Color.White;
			this.ClientSize = new System.Drawing.Size(1280, 961);
			this.Controls.Add(this.panel1);
			this.Controls.Add(this.tableLayoutPanel1);
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.MaximumSize = new System.Drawing.Size(1296, 1000);
			this.MinimumSize = new System.Drawing.Size(1278, 982);
			this.Name = "frmHypMoldLineInformation";
			this.Text = "HYPOID PINION";
			this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
			this.Load += new System.EventHandler(this.Form1_Load);
			this.tableLayoutPanel1.ResumeLayout(false);
			this.panel2.ResumeLayout(false);
			this.panel2.PerformLayout();
			this.panel3.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.grdData)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grvData)).EndInit();
			this.panel1.ResumeLayout(false);
			this.panel1.PerformLayout();
			this.mnuMenu.ResumeLayout(false);
			this.mnuMenu.PerformLayout();
			this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
		private System.Windows.Forms.Panel panel1;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.ToolStrip mnuMenu;
		private System.Windows.Forms.ToolStripButton btnSave;
		private System.Windows.Forms.ToolStripButton btnPath;
		private System.Windows.Forms.TextBox txtMachine;
		private System.Windows.Forms.ComboBox cboCongDoan;
		private System.Windows.Forms.TextBox txtOrder;
		private System.Windows.Forms.TextBox txtKnifeDetailCode;
		private System.Windows.Forms.TextBox txtWorkerCode;
		private System.Windows.Forms.Label lbly;
		private System.Windows.Forms.Label lblx;
		private DevExpress.XtraPdfViewer.PdfViewer pdfViewer1;
		private System.Windows.Forms.Panel panel2;
		private System.Windows.Forms.Panel panel3;
		private DevExpress.XtraGrid.GridControl grdData;
		private DevExpress.XtraGrid.Views.Grid.GridView grvData;
		private DevExpress.XtraGrid.Columns.GridColumn colWorkingName;
		private DevExpress.XtraGrid.Columns.GridColumn colPeriodValue;
		private DevExpress.XtraGrid.Columns.GridColumn colRealvalue;
		private DevExpress.XtraGrid.Columns.GridColumn colSortOrder;
		private DevExpress.XtraGrid.Columns.GridColumn colCondition;
		private DevExpress.XtraGrid.Columns.GridColumn colUnit;
		private DevExpress.XtraGrid.Columns.GridColumn colResult;
		private DevExpress.XtraGrid.Columns.GridColumn colX;
		private DevExpress.XtraGrid.Columns.GridColumn colY;
		private DevExpress.XtraGrid.Columns.GridColumn colMinValue;
		private DevExpress.XtraGrid.Columns.GridColumn colMaxValue;
		private DevExpress.XtraGrid.Columns.GridColumn colisDisplay;
		private DevExpress.XtraGrid.Columns.GridColumn colCheckValueType;
		private DevExpress.XtraGrid.Columns.GridColumn colID;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.TextBox txtGoodsCode;
		private System.Windows.Forms.ComboBox cboPart;
	}
}

