
namespace BMS
{
    partial class frmJigLineInformation
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
			DevExpress.XtraCharts.XYDiagram xyDiagram1 = new DevExpress.XtraCharts.XYDiagram();
			DevExpress.XtraCharts.SecondaryAxisY secondaryAxisY1 = new DevExpress.XtraCharts.SecondaryAxisY();
			DevExpress.XtraCharts.Series series1 = new DevExpress.XtraCharts.Series();
			DevExpress.XtraCharts.PointSeriesLabel pointSeriesLabel1 = new DevExpress.XtraCharts.PointSeriesLabel();
			DevExpress.XtraCharts.SeriesPoint seriesPoint1 = new DevExpress.XtraCharts.SeriesPoint("09/08/2021", new object[] {
            ((object)(0.01D))});
			DevExpress.XtraCharts.SeriesPoint seriesPoint2 = new DevExpress.XtraCharts.SeriesPoint("10/08/2021", new object[] {
            ((object)(0.1D))});
			DevExpress.XtraCharts.LineSeriesView lineSeriesView1 = new DevExpress.XtraCharts.LineSeriesView();
			DevExpress.XtraCharts.Series series2 = new DevExpress.XtraCharts.Series();
			DevExpress.XtraCharts.SeriesPoint seriesPoint3 = new DevExpress.XtraCharts.SeriesPoint("09/08/2021", new object[] {
            ((object)(0.1D))});
			DevExpress.XtraCharts.SeriesPoint seriesPoint4 = new DevExpress.XtraCharts.SeriesPoint("10/08/2021", new object[] {
            ((object)(0.1D))});
			DevExpress.XtraCharts.LineSeriesView lineSeriesView2 = new DevExpress.XtraCharts.LineSeriesView();
			DevExpress.XtraCharts.Series series3 = new DevExpress.XtraCharts.Series();
			DevExpress.XtraCharts.SeriesPoint seriesPoint5 = new DevExpress.XtraCharts.SeriesPoint("09/08/2021", new object[] {
            ((object)(-0.1D))});
			DevExpress.XtraCharts.SeriesPoint seriesPoint6 = new DevExpress.XtraCharts.SeriesPoint("10/08/2021", new object[] {
            ((object)(-0.1D))});
			DevExpress.XtraCharts.LineSeriesView lineSeriesView3 = new DevExpress.XtraCharts.LineSeriesView();
			DevExpress.XtraCharts.LineSeriesView lineSeriesView4 = new DevExpress.XtraCharts.LineSeriesView();
			this.repositoryItemMemoEdit2 = new DevExpress.XtraEditors.Repository.RepositoryItemMemoEdit();
			this.txtDateTime = new System.Windows.Forms.TextBox();
			this.txtMachine = new System.Windows.Forms.TextBox();
			this.cboStep = new System.Windows.Forms.ComboBox();
			this.txtJigCode = new System.Windows.Forms.TextBox();
			this.txtWorker = new System.Windows.Forms.TextBox();
			this.grdData = new DevExpress.XtraGrid.GridControl();
			this.grvData = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.colSortOrder = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colProductWorkingName = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colStandardValue = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colMinValue = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colMaxValue = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colReal = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colRate = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colCheckValueType = new DevExpress.XtraGrid.Columns.GridColumn();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.pannelBot = new System.Windows.Forms.Panel();
			this.btnSave = new System.Windows.Forms.Button();
			this.label7 = new System.Windows.Forms.Label();
			this.txtGoods = new System.Windows.Forms.TextBox();
			this.cbPart = new System.Windows.Forms.ComboBox();
			this.label8 = new System.Windows.Forms.Label();
			this.txtOrder = new System.Windows.Forms.TextBox();
			this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
			this.chartControl1 = new DevExpress.XtraCharts.ChartControl();
			this.tableLayoutPanel2 = new System.Windows.Forms.TableLayoutPanel();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemMemoEdit2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grdData)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grvData)).BeginInit();
			this.pannelBot.SuspendLayout();
			this.tableLayoutPanel1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.chartControl1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(xyDiagram1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(secondaryAxisY1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(series1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(pointSeriesLabel1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(lineSeriesView1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(series2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(lineSeriesView2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(series3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(lineSeriesView3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(lineSeriesView4)).BeginInit();
			this.tableLayoutPanel2.SuspendLayout();
			this.SuspendLayout();
			// 
			// repositoryItemMemoEdit2
			// 
			this.repositoryItemMemoEdit2.Appearance.Options.UseTextOptions = true;
			this.repositoryItemMemoEdit2.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.repositoryItemMemoEdit2.Appearance.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.repositoryItemMemoEdit2.Appearance.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.repositoryItemMemoEdit2.Name = "repositoryItemMemoEdit2";
			// 
			// txtDateTime
			// 
			this.txtDateTime.Dock = System.Windows.Forms.DockStyle.Fill;
			this.txtDateTime.Font = new System.Drawing.Font("Microsoft Sans Serif", 21.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtDateTime.Location = new System.Drawing.Point(1029, 1);
			this.txtDateTime.Margin = new System.Windows.Forms.Padding(1);
			this.txtDateTime.Name = "txtDateTime";
			this.txtDateTime.Size = new System.Drawing.Size(224, 40);
			this.txtDateTime.TabIndex = 85;
			// 
			// txtMachine
			// 
			this.txtMachine.Dock = System.Windows.Forms.DockStyle.Fill;
			this.txtMachine.Font = new System.Drawing.Font("Microsoft Sans Serif", 21.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtMachine.Location = new System.Drawing.Point(1029, 42);
			this.txtMachine.Margin = new System.Windows.Forms.Padding(1);
			this.txtMachine.Name = "txtMachine";
			this.txtMachine.Size = new System.Drawing.Size(224, 40);
			this.txtMachine.TabIndex = 0;
			// 
			// cboStep
			// 
			this.cboStep.Dock = System.Windows.Forms.DockStyle.Fill;
			this.cboStep.Font = new System.Drawing.Font("Microsoft Sans Serif", 21.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.cboStep.FormattingEnabled = true;
			this.cboStep.Location = new System.Drawing.Point(735, 42);
			this.cboStep.Margin = new System.Windows.Forms.Padding(1);
			this.cboStep.Name = "cboStep";
			this.cboStep.Size = new System.Drawing.Size(193, 41);
			this.cboStep.TabIndex = 5;
			// 
			// txtJigCode
			// 
			this.txtJigCode.Dock = System.Windows.Forms.DockStyle.Fill;
			this.txtJigCode.Font = new System.Drawing.Font("Microsoft Sans Serif", 21.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtJigCode.Location = new System.Drawing.Point(428, 42);
			this.txtJigCode.Margin = new System.Windows.Forms.Padding(1);
			this.txtJigCode.Name = "txtJigCode";
			this.txtJigCode.Size = new System.Drawing.Size(191, 40);
			this.txtJigCode.TabIndex = 4;
			// 
			// txtWorker
			// 
			this.txtWorker.Dock = System.Windows.Forms.DockStyle.Fill;
			this.txtWorker.Font = new System.Drawing.Font("Microsoft Sans Serif", 21.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtWorker.Location = new System.Drawing.Point(428, 1);
			this.txtWorker.Margin = new System.Windows.Forms.Padding(1);
			this.txtWorker.Name = "txtWorker";
			this.txtWorker.Size = new System.Drawing.Size(191, 40);
			this.txtWorker.TabIndex = 2;
			// 
			// grdData
			// 
			this.grdData.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.grdData.Font = new System.Drawing.Font("Times New Roman", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.grdData.Location = new System.Drawing.Point(3, 3);
			this.grdData.MainView = this.grvData;
			this.grdData.Name = "grdData";
			this.grdData.Size = new System.Drawing.Size(779, 428);
			this.grdData.TabIndex = 6;
			this.grdData.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.grvData});
			// 
			// grvData
			// 
			this.grvData.Appearance.FocusedCell.BorderColor = System.Drawing.Color.Lime;
			this.grvData.Appearance.FocusedCell.Options.UseBorderColor = true;
			this.grvData.Appearance.HeaderPanel.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(51)))), ((int)(((byte)(204)))));
			this.grvData.Appearance.HeaderPanel.ForeColor = System.Drawing.Color.White;
			this.grvData.Appearance.HeaderPanel.Options.UseBackColor = true;
			this.grvData.Appearance.HeaderPanel.Options.UseBorderColor = true;
			this.grvData.Appearance.HeaderPanel.Options.UseFont = true;
			this.grvData.Appearance.HeaderPanel.Options.UseForeColor = true;
			this.grvData.Appearance.HeaderPanel.Options.UseImage = true;
			this.grvData.Appearance.HeaderPanel.Options.UseTextOptions = true;
			this.grvData.Appearance.HeaderPanel.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.grvData.Appearance.HeaderPanel.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.grvData.Appearance.HeaderPanel.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.grvData.Appearance.ViewCaption.Options.UseTextOptions = true;
			this.grvData.Appearance.ViewCaption.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.grvData.ColumnPanelRowHeight = 50;
			this.grvData.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
            this.colSortOrder,
            this.colProductWorkingName,
            this.colStandardValue,
            this.colMinValue,
            this.colMaxValue,
            this.colReal,
            this.colRate,
            this.colCheckValueType});
			this.grvData.GridControl = this.grdData;
			this.grvData.HorzScrollStep = 5;
			this.grvData.Name = "grvData";
			this.grvData.OptionsBehavior.AutoExpandAllGroups = true;
			this.grvData.OptionsCustomization.AllowFilter = false;
			this.grvData.OptionsCustomization.AllowSort = false;
			this.grvData.OptionsFilter.AllowFilterEditor = false;
			this.grvData.OptionsFind.ShowCloseButton = false;
			this.grvData.OptionsLayout.Columns.AddNewColumns = false;
			this.grvData.OptionsLayout.Columns.RemoveOldColumns = false;
			this.grvData.OptionsSelection.EnableAppearanceFocusedCell = false;
			this.grvData.OptionsSelection.EnableAppearanceFocusedRow = false;
			this.grvData.OptionsView.RowAutoHeight = true;
			this.grvData.OptionsView.ShowGroupPanel = false;
			this.grvData.PaintStyleName = "Flat";
			this.grvData.Click += new System.EventHandler(this.grvData_Click);
			// 
			// colSortOrder
			// 
			this.colSortOrder.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.colSortOrder.AppearanceCell.Options.UseFont = true;
			this.colSortOrder.AppearanceCell.Options.UseTextOptions = true;
			this.colSortOrder.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colSortOrder.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colSortOrder.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colSortOrder.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 16F, System.Drawing.FontStyle.Bold);
			this.colSortOrder.AppearanceHeader.Options.UseFont = true;
			this.colSortOrder.AppearanceHeader.Options.UseTextOptions = true;
			this.colSortOrder.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colSortOrder.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colSortOrder.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colSortOrder.Caption = "STT";
			this.colSortOrder.FieldName = "SortOrder";
			this.colSortOrder.Name = "colSortOrder";
			this.colSortOrder.OptionsColumn.AllowEdit = false;
			this.colSortOrder.OptionsColumn.AllowFocus = false;
			this.colSortOrder.OptionsColumn.AllowSize = false;
			this.colSortOrder.OptionsColumn.ReadOnly = true;
			this.colSortOrder.Visible = true;
			this.colSortOrder.VisibleIndex = 0;
			// 
			// colProductWorkingName
			// 
			this.colProductWorkingName.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.colProductWorkingName.AppearanceCell.Options.UseFont = true;
			this.colProductWorkingName.AppearanceCell.Options.UseTextOptions = true;
			this.colProductWorkingName.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colProductWorkingName.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(128)))), ((int)(((byte)(255)))));
			this.colProductWorkingName.AppearanceHeader.BackColor2 = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(128)))), ((int)(((byte)(255)))));
			this.colProductWorkingName.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 16F, System.Drawing.FontStyle.Bold);
			this.colProductWorkingName.AppearanceHeader.Options.UseFont = true;
			this.colProductWorkingName.AppearanceHeader.Options.UseTextOptions = true;
			this.colProductWorkingName.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colProductWorkingName.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colProductWorkingName.Caption = "Mục kiểm tra";
			this.colProductWorkingName.FieldName = "WorkingName";
			this.colProductWorkingName.Name = "colProductWorkingName";
			this.colProductWorkingName.OptionsColumn.AllowEdit = false;
			this.colProductWorkingName.OptionsColumn.AllowFocus = false;
			this.colProductWorkingName.OptionsColumn.ReadOnly = true;
			this.colProductWorkingName.Visible = true;
			this.colProductWorkingName.VisibleIndex = 1;
			this.colProductWorkingName.Width = 494;
			// 
			// colStandardValue
			// 
			this.colStandardValue.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.colStandardValue.AppearanceCell.Options.UseFont = true;
			this.colStandardValue.AppearanceCell.Options.UseTextOptions = true;
			this.colStandardValue.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colStandardValue.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colStandardValue.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colStandardValue.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colStandardValue.AppearanceHeader.Options.UseFont = true;
			this.colStandardValue.AppearanceHeader.Options.UseTextOptions = true;
			this.colStandardValue.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colStandardValue.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colStandardValue.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colStandardValue.Caption = "Giá trị tiêu chuẩn";
			this.colStandardValue.ColumnEdit = this.repositoryItemMemoEdit2;
			this.colStandardValue.FieldName = "StandardValue";
			this.colStandardValue.Name = "colStandardValue";
			this.colStandardValue.OptionsColumn.AllowEdit = false;
			this.colStandardValue.OptionsColumn.AllowFocus = false;
			this.colStandardValue.OptionsColumn.AllowSize = false;
			this.colStandardValue.OptionsColumn.ReadOnly = true;
			this.colStandardValue.Visible = true;
			this.colStandardValue.VisibleIndex = 2;
			this.colStandardValue.Width = 263;
			// 
			// colMinValue
			// 
			this.colMinValue.AppearanceCell.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.colMinValue.AppearanceCell.Options.UseFont = true;
			this.colMinValue.AppearanceCell.Options.UseTextOptions = true;
			this.colMinValue.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colMinValue.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colMinValue.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colMinValue.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colMinValue.AppearanceHeader.Options.UseFont = true;
			this.colMinValue.AppearanceHeader.Options.UseTextOptions = true;
			this.colMinValue.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colMinValue.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colMinValue.Caption = "MinValue";
			this.colMinValue.FieldName = "MinValue";
			this.colMinValue.Name = "colMinValue";
			this.colMinValue.OptionsColumn.AllowEdit = false;
			this.colMinValue.OptionsColumn.ReadOnly = true;
			this.colMinValue.Width = 104;
			// 
			// colMaxValue
			// 
			this.colMaxValue.AppearanceCell.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.colMaxValue.AppearanceCell.Options.UseFont = true;
			this.colMaxValue.AppearanceCell.Options.UseTextOptions = true;
			this.colMaxValue.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colMaxValue.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colMaxValue.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colMaxValue.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colMaxValue.AppearanceHeader.Options.UseFont = true;
			this.colMaxValue.AppearanceHeader.Options.UseTextOptions = true;
			this.colMaxValue.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colMaxValue.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colMaxValue.Caption = "MaxValue";
			this.colMaxValue.FieldName = "MaxValue";
			this.colMaxValue.Name = "colMaxValue";
			this.colMaxValue.OptionsColumn.AllowEdit = false;
			this.colMaxValue.OptionsColumn.ReadOnly = true;
			// 
			// colReal
			// 
			this.colReal.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.colReal.AppearanceCell.Options.UseFont = true;
			this.colReal.AppearanceCell.Options.UseTextOptions = true;
			this.colReal.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colReal.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colReal.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colReal.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colReal.AppearanceHeader.Options.UseFont = true;
			this.colReal.AppearanceHeader.Options.UseTextOptions = true;
			this.colReal.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colReal.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colReal.Caption = "Thực tế";
			this.colReal.FieldName = "RealValue";
			this.colReal.Name = "colReal";
			this.colReal.OptionsColumn.AllowEdit = false;
			this.colReal.OptionsColumn.AllowSize = false;
			this.colReal.OptionsColumn.ReadOnly = true;
			this.colReal.Visible = true;
			this.colReal.VisibleIndex = 3;
			this.colReal.Width = 118;
			// 
			// colRate
			// 
			this.colRate.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.colRate.AppearanceCell.Options.UseFont = true;
			this.colRate.AppearanceCell.Options.UseTextOptions = true;
			this.colRate.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colRate.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colRate.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colRate.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colRate.AppearanceHeader.Options.UseFont = true;
			this.colRate.AppearanceHeader.Options.UseTextOptions = true;
			this.colRate.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colRate.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colRate.Caption = "Đánh giá";
			this.colRate.FieldName = "Result";
			this.colRate.Name = "colRate";
			this.colRate.OptionsColumn.AllowEdit = false;
			this.colRate.OptionsColumn.AllowFocus = false;
			this.colRate.OptionsColumn.AllowSize = false;
			this.colRate.OptionsColumn.ReadOnly = true;
			this.colRate.Visible = true;
			this.colRate.VisibleIndex = 4;
			this.colRate.Width = 109;
			// 
			// colCheckValueType
			// 
			this.colCheckValueType.Caption = "CheckValueType";
			this.colCheckValueType.FieldName = "CheckValueType";
			this.colCheckValueType.Name = "colCheckValueType";
			this.colCheckValueType.OptionsColumn.AllowEdit = false;
			this.colCheckValueType.OptionsColumn.ReadOnly = true;
			// 
			// label1
			// 
			this.label1.BackColor = System.Drawing.Color.RoyalBlue;
			this.label1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold);
			this.label1.ForeColor = System.Drawing.Color.White;
			this.label1.Location = new System.Drawing.Point(311, 1);
			this.label1.Margin = new System.Windows.Forms.Padding(1);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(115, 39);
			this.label1.TabIndex = 90;
			this.label1.Text = "PHỤ TRÁCH";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label2
			// 
			this.label2.BackColor = System.Drawing.Color.RoyalBlue;
			this.label2.Dock = System.Windows.Forms.DockStyle.Fill;
			this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold);
			this.label2.ForeColor = System.Drawing.Color.White;
			this.label2.Location = new System.Drawing.Point(311, 42);
			this.label2.Margin = new System.Windows.Forms.Padding(1);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(115, 40);
			this.label2.TabIndex = 91;
			this.label2.Text = "MÃ DỤNG CỤ";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label3
			// 
			this.label3.BackColor = System.Drawing.Color.RoyalBlue;
			this.label3.Dock = System.Windows.Forms.DockStyle.Fill;
			this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold);
			this.label3.ForeColor = System.Drawing.Color.White;
			this.label3.Location = new System.Drawing.Point(621, 1);
			this.label3.Margin = new System.Windows.Forms.Padding(1);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(112, 39);
			this.label3.TabIndex = 92;
			this.label3.Text = "BỘ PHẬN";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label4
			// 
			this.label4.BackColor = System.Drawing.Color.RoyalBlue;
			this.label4.Dock = System.Windows.Forms.DockStyle.Fill;
			this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold);
			this.label4.ForeColor = System.Drawing.Color.White;
			this.label4.Location = new System.Drawing.Point(930, 42);
			this.label4.Margin = new System.Windows.Forms.Padding(1);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(97, 40);
			this.label4.TabIndex = 93;
			this.label4.Text = "TÊN MÁY";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label5
			// 
			this.label5.BackColor = System.Drawing.Color.RoyalBlue;
			this.label5.Dock = System.Windows.Forms.DockStyle.Fill;
			this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold);
			this.label5.ForeColor = System.Drawing.Color.White;
			this.label5.Location = new System.Drawing.Point(621, 42);
			this.label5.Margin = new System.Windows.Forms.Padding(1);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(112, 40);
			this.label5.TabIndex = 94;
			this.label5.Text = "CÔNG ĐOẠN";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label6
			// 
			this.label6.BackColor = System.Drawing.Color.RoyalBlue;
			this.label6.Dock = System.Windows.Forms.DockStyle.Fill;
			this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold);
			this.label6.ForeColor = System.Drawing.Color.White;
			this.label6.Location = new System.Drawing.Point(930, 1);
			this.label6.Margin = new System.Windows.Forms.Padding(1);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(97, 39);
			this.label6.TabIndex = 95;
			this.label6.Text = "THỜI GIAN";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// pannelBot
			// 
			this.pannelBot.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.pannelBot.Controls.Add(this.btnSave);
			this.pannelBot.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
			this.pannelBot.Location = new System.Drawing.Point(0, 522);
			this.pannelBot.Margin = new System.Windows.Forms.Padding(2);
			this.pannelBot.Name = "pannelBot";
			this.pannelBot.Size = new System.Drawing.Size(1254, 77);
			this.pannelBot.TabIndex = 8;
			this.pannelBot.Visible = false;
			// 
			// btnSave
			// 
			this.btnSave.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnSave.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnSave.Location = new System.Drawing.Point(1095, 4);
			this.btnSave.Name = "btnSave";
			this.btnSave.Size = new System.Drawing.Size(154, 33);
			this.btnSave.TabIndex = 0;
			this.btnSave.Text = "CẤT DỮ LIỆU";
			this.btnSave.UseVisualStyleBackColor = true;
			this.btnSave.Visible = false;
			// 
			// label7
			// 
			this.label7.BackColor = System.Drawing.Color.RoyalBlue;
			this.label7.Dock = System.Windows.Forms.DockStyle.Fill;
			this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.label7.ForeColor = System.Drawing.Color.White;
			this.label7.Location = new System.Drawing.Point(1, 1);
			this.label7.Margin = new System.Windows.Forms.Padding(1);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(106, 39);
			this.label7.TabIndex = 98;
			this.label7.Text = "MÃ HÀNG";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// txtGoods
			// 
			this.txtGoods.Dock = System.Windows.Forms.DockStyle.Fill;
			this.txtGoods.Font = new System.Drawing.Font("Microsoft Sans Serif", 21.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtGoods.Location = new System.Drawing.Point(109, 1);
			this.txtGoods.Margin = new System.Windows.Forms.Padding(1);
			this.txtGoods.Name = "txtGoods";
			this.txtGoods.Size = new System.Drawing.Size(200, 40);
			this.txtGoods.TabIndex = 3;
			// 
			// cbPart
			// 
			this.cbPart.Dock = System.Windows.Forms.DockStyle.Fill;
			this.cbPart.Font = new System.Drawing.Font("Microsoft Sans Serif", 21.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.cbPart.FormattingEnabled = true;
			this.cbPart.Location = new System.Drawing.Point(735, 1);
			this.cbPart.Margin = new System.Windows.Forms.Padding(1);
			this.cbPart.Name = "cbPart";
			this.cbPart.Size = new System.Drawing.Size(193, 41);
			this.cbPart.TabIndex = 99;
			// 
			// label8
			// 
			this.label8.BackColor = System.Drawing.Color.RoyalBlue;
			this.label8.Dock = System.Windows.Forms.DockStyle.Fill;
			this.label8.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold);
			this.label8.ForeColor = System.Drawing.Color.White;
			this.label8.Location = new System.Drawing.Point(1, 42);
			this.label8.Margin = new System.Windows.Forms.Padding(1);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(106, 40);
			this.label8.TabIndex = 101;
			this.label8.Text = "ORDER";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// txtOrder
			// 
			this.txtOrder.Dock = System.Windows.Forms.DockStyle.Fill;
			this.txtOrder.Font = new System.Drawing.Font("Microsoft Sans Serif", 21.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtOrder.Location = new System.Drawing.Point(109, 42);
			this.txtOrder.Margin = new System.Windows.Forms.Padding(1);
			this.txtOrder.Name = "txtOrder";
			this.txtOrder.Size = new System.Drawing.Size(200, 40);
			this.txtOrder.TabIndex = 100;
			// 
			// tableLayoutPanel1
			// 
			this.tableLayoutPanel1.ColumnCount = 8;
			this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 8.66426F));
			this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 16.24549F));
			this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 9.386282F));
			this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 15.52347F));
			this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 9.205776F));
			this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 15.70397F));
			this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 7.942238F));
			this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 17.87004F));
			this.tableLayoutPanel1.Controls.Add(this.label7, 0, 0);
			this.tableLayoutPanel1.Controls.Add(this.cbPart, 5, 0);
			this.tableLayoutPanel1.Controls.Add(this.txtWorker, 3, 0);
			this.tableLayoutPanel1.Controls.Add(this.txtMachine, 7, 1);
			this.tableLayoutPanel1.Controls.Add(this.txtJigCode, 3, 1);
			this.tableLayoutPanel1.Controls.Add(this.txtDateTime, 7, 0);
			this.tableLayoutPanel1.Controls.Add(this.label4, 6, 1);
			this.tableLayoutPanel1.Controls.Add(this.label6, 6, 0);
			this.tableLayoutPanel1.Controls.Add(this.txtOrder, 1, 1);
			this.tableLayoutPanel1.Controls.Add(this.label8, 0, 1);
			this.tableLayoutPanel1.Controls.Add(this.txtGoods, 1, 0);
			this.tableLayoutPanel1.Controls.Add(this.label5, 4, 1);
			this.tableLayoutPanel1.Controls.Add(this.cboStep, 5, 1);
			this.tableLayoutPanel1.Controls.Add(this.label2, 2, 1);
			this.tableLayoutPanel1.Controls.Add(this.label3, 4, 0);
			this.tableLayoutPanel1.Controls.Add(this.label1, 2, 0);
			this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Top;
			this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
			this.tableLayoutPanel1.Name = "tableLayoutPanel1";
			this.tableLayoutPanel1.RowCount = 2;
			this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
			this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
			this.tableLayoutPanel1.Size = new System.Drawing.Size(1254, 83);
			this.tableLayoutPanel1.TabIndex = 102;
			// 
			// chartControl1
			// 
			this.chartControl1.AppearanceNameSerializable = "Chameleon";
			this.chartControl1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
			this.chartControl1.BorderOptions.Color = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
			this.chartControl1.BorderOptions.Visible = false;
			xyDiagram1.AxisX.Color = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
			xyDiagram1.AxisX.Label.TextColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
			xyDiagram1.AxisX.Title.Font = new System.Drawing.Font("Tahoma", 10F);
			xyDiagram1.AxisX.Title.Text = "Ngày/Tháng";
			xyDiagram1.AxisX.VisibleInPanesSerializable = "-1";
			xyDiagram1.AxisX.WholeRange.AutoSideMargins = true;
			xyDiagram1.AxisY.Color = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
			xyDiagram1.AxisY.Label.TextColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
			xyDiagram1.AxisY.Title.Font = new System.Drawing.Font("Tahoma", 10F);
			xyDiagram1.AxisY.Title.Text = "";
			xyDiagram1.AxisY.Title.TextColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
			xyDiagram1.AxisY.Title.Visible = true;
			xyDiagram1.AxisY.VisibleInPanesSerializable = "-1";
			xyDiagram1.DefaultPane.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
			xyDiagram1.DefaultPane.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
			xyDiagram1.DefaultPane.FillStyle.FillMode = DevExpress.XtraCharts.FillMode.Solid;
			xyDiagram1.Margins.Bottom = 13;
			xyDiagram1.Margins.Left = 13;
			xyDiagram1.Margins.Right = 13;
			xyDiagram1.Margins.Top = 13;
			xyDiagram1.PaneDistance = 18;
			secondaryAxisY1.AxisID = 0;
			secondaryAxisY1.Color = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
			secondaryAxisY1.Label.TextColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
			secondaryAxisY1.Name = "Secondary AxisY 1";
			secondaryAxisY1.Title.Font = new System.Drawing.Font("Tahoma", 10F);
			secondaryAxisY1.Title.Text = "";
			secondaryAxisY1.Title.TextColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
			secondaryAxisY1.Title.Visible = true;
			secondaryAxisY1.VisibleInPanesSerializable = "-1";
			xyDiagram1.SecondaryAxesY.AddRange(new DevExpress.XtraCharts.SecondaryAxisY[] {
            secondaryAxisY1});
			this.chartControl1.Diagram = xyDiagram1;
			this.chartControl1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.chartControl1.EmptyChartText.Text = "Không tải được dữ liêu\r\n";
			this.chartControl1.Legend.AlignmentHorizontal = DevExpress.XtraCharts.LegendAlignmentHorizontal.Center;
			this.chartControl1.Legend.AlignmentVertical = DevExpress.XtraCharts.LegendAlignmentVertical.TopOutside;
			this.chartControl1.Legend.Direction = DevExpress.XtraCharts.LegendDirection.LeftToRight;
			this.chartControl1.Legend.EquallySpacedItems = false;
			this.chartControl1.Legend.Margins.Bottom = 0;
			this.chartControl1.Legend.Margins.Left = 0;
			this.chartControl1.Legend.Margins.Right = 0;
			this.chartControl1.Legend.Margins.Top = 0;
			this.chartControl1.Legend.UseCheckBoxes = true;
			this.chartControl1.Location = new System.Drawing.Point(788, 3);
			this.chartControl1.Name = "chartControl1";
			this.chartControl1.Padding.Bottom = 4;
			this.chartControl1.Padding.Left = 4;
			this.chartControl1.Padding.Right = 4;
			this.chartControl1.Padding.Top = 4;
			this.chartControl1.PaletteName = "Palette 1";
			this.chartControl1.PaletteRepository.Add("Palette 1", new DevExpress.XtraCharts.Palette("Palette 1", DevExpress.XtraCharts.PaletteScaleMode.Repeat, new DevExpress.XtraCharts.PaletteEntry[] {
                new DevExpress.XtraCharts.PaletteEntry(System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(206)))), ((int)(((byte)(52))))), System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(206)))), ((int)(((byte)(52)))))),
                new DevExpress.XtraCharts.PaletteEntry(System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(202)))), ((int)(((byte)(202))))), System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(202)))), ((int)(((byte)(202)))))),
                new DevExpress.XtraCharts.PaletteEntry(System.Drawing.Color.Red, System.Drawing.Color.Red)}));
			this.chartControl1.RuntimeHitTesting = true;
			this.chartControl1.RuntimeSeriesSelectionMode = DevExpress.XtraCharts.SeriesSelectionMode.Point;
			series1.CheckableInLegend = false;
			pointSeriesLabel1.Angle = 90;
			pointSeriesLabel1.Antialiasing = true;
			pointSeriesLabel1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
			pointSeriesLabel1.LineColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
			pointSeriesLabel1.LineStyle.DashStyle = DevExpress.XtraCharts.DashStyle.Dot;
			pointSeriesLabel1.LineStyle.Thickness = 2;
			pointSeriesLabel1.LineVisible = false;
			pointSeriesLabel1.ResolveOverlappingMode = DevExpress.XtraCharts.ResolveOverlappingMode.HideOverlapped;
			series1.Label = pointSeriesLabel1;
			series1.LabelsVisibility = DevExpress.Utils.DefaultBoolean.True;
			series1.Name = "Thực tế";
			series1.Points.AddRange(new DevExpress.XtraCharts.SeriesPoint[] {
            seriesPoint1,
            seriesPoint2});
			lineSeriesView1.LineMarkerOptions.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
			lineSeriesView1.LineMarkerOptions.Color = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
			lineSeriesView1.LineMarkerOptions.FillStyle.FillMode = DevExpress.XtraCharts.FillMode.Hatch;
			lineSeriesView1.LineMarkerOptions.Size = 7;
			lineSeriesView1.MarkerVisibility = DevExpress.Utils.DefaultBoolean.True;
			series1.View = lineSeriesView1;
			series2.CheckableInLegend = false;
			series2.Name = "Max";
			series2.Points.AddRange(new DevExpress.XtraCharts.SeriesPoint[] {
            seriesPoint3,
            seriesPoint4});
			series2.View = lineSeriesView2;
			series3.CheckableInLegend = false;
			series3.Name = "Min";
			series3.Points.AddRange(new DevExpress.XtraCharts.SeriesPoint[] {
            seriesPoint5,
            seriesPoint6});
			series3.View = lineSeriesView3;
			this.chartControl1.SeriesSerializable = new DevExpress.XtraCharts.Series[] {
        series1,
        series2,
        series3};
			this.chartControl1.SeriesTemplate.View = lineSeriesView4;
			this.chartControl1.SideBySideBarDistanceVariable = 0.1D;
			this.chartControl1.Size = new System.Drawing.Size(462, 428);
			this.chartControl1.TabIndex = 223;
			// 
			// tableLayoutPanel2
			// 
			this.tableLayoutPanel2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.tableLayoutPanel2.ColumnCount = 2;
			this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 62.68657F));
			this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 37.31343F));
			this.tableLayoutPanel2.Controls.Add(this.chartControl1, 1, 0);
			this.tableLayoutPanel2.Controls.Add(this.grdData, 0, 0);
			this.tableLayoutPanel2.Location = new System.Drawing.Point(0, 86);
			this.tableLayoutPanel2.Name = "tableLayoutPanel2";
			this.tableLayoutPanel2.RowCount = 1;
			this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
			this.tableLayoutPanel2.Size = new System.Drawing.Size(1253, 434);
			this.tableLayoutPanel2.TabIndex = 224;
			// 
			// frmJigLineInformation
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.BackColor = System.Drawing.Color.White;
			this.ClientSize = new System.Drawing.Size(1254, 599);
			this.Controls.Add(this.tableLayoutPanel2);
			this.Controls.Add(this.tableLayoutPanel1);
			this.Controls.Add(this.pannelBot);
			this.Name = "frmJigLineInformation";
			this.Text = "LINE JIG";
			this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
			this.Load += new System.EventHandler(this.frmDaoJigLine_Load);
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemMemoEdit2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grdData)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grvData)).EndInit();
			this.pannelBot.ResumeLayout(false);
			this.tableLayoutPanel1.ResumeLayout(false);
			this.tableLayoutPanel1.PerformLayout();
			((System.ComponentModel.ISupportInitialize)(secondaryAxisY1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(xyDiagram1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(pointSeriesLabel1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(lineSeriesView1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(series1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(lineSeriesView2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(series2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(lineSeriesView3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(series3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(lineSeriesView4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chartControl1)).EndInit();
			this.tableLayoutPanel2.ResumeLayout(false);
			this.ResumeLayout(false);

        }

        #endregion
		private System.Windows.Forms.TextBox txtDateTime;
		private System.Windows.Forms.TextBox txtMachine;
		private System.Windows.Forms.ComboBox cboStep;
		private System.Windows.Forms.TextBox txtJigCode;
		private System.Windows.Forms.TextBox txtWorker;
		private DevExpress.XtraGrid.GridControl grdData;
		private DevExpress.XtraGrid.Views.Grid.GridView grvData;
		private DevExpress.XtraGrid.Columns.GridColumn colProductWorkingName;
		private DevExpress.XtraGrid.Columns.GridColumn colStandardValue;
		private DevExpress.XtraGrid.Columns.GridColumn colMinValue;
		private DevExpress.XtraGrid.Columns.GridColumn colMaxValue;
		private DevExpress.XtraGrid.Columns.GridColumn colSortOrder;
		private DevExpress.XtraGrid.Columns.GridColumn colReal;
		private DevExpress.XtraGrid.Columns.GridColumn colRate;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Panel pannelBot;
		private DevExpress.XtraGrid.Columns.GridColumn colCheckValueType;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.TextBox txtGoods;
		private DevExpress.XtraEditors.Repository.RepositoryItemMemoEdit repositoryItemMemoEdit2;
		private System.Windows.Forms.ComboBox cbPart;
		private System.Windows.Forms.Button btnSave;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.TextBox txtOrder;
		private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
		private DevExpress.XtraCharts.ChartControl chartControl1;
		private System.Windows.Forms.TableLayoutPanel tableLayoutPanel2;
	}
}

