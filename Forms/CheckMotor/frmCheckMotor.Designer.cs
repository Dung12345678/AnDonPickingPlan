
namespace BMS
{
    partial class frmCheckMotor
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
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmCheckMotor));
			this.toolStrip1 = new System.Windows.Forms.ToolStrip();
			this.btnCreatePlan = new System.Windows.Forms.ToolStripButton();
			this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
			this.btnEditPlan = new System.Windows.Forms.ToolStripButton();
			this.toolStripSeparator7 = new System.Windows.Forms.ToolStripSeparator();
			this.btnDelPlan = new System.Windows.Forms.ToolStripButton();
			this.grdData = new DevExpress.XtraGrid.GridControl();
			this.grvData = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.colID = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colSalesOrder = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colOrderedQty = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colQtyOfShipOrder = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colCardNo = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colMotorInspSealNo = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colSerialNo = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colArticleID = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colAssemblyOrderNo = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colNoOfJG = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colDescriptions = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colUPR = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colJGDate = new DevExpress.XtraGrid.Columns.GridColumn();
			this.timer1 = new System.Windows.Forms.Timer(this.components);
			this.panel2 = new System.Windows.Forms.Panel();
			this.btnSearch = new DevExpress.XtraEditors.SimpleButton();
			this.txbSearch = new System.Windows.Forms.TextBox();
			this.dtpFrom = new System.Windows.Forms.DateTimePicker();
			this.dtpTo = new System.Windows.Forms.DateTimePicker();
			this.label1 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.toolStrip1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.grdData)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grvData)).BeginInit();
			this.panel2.SuspendLayout();
			this.SuspendLayout();
			// 
			// toolStrip1
			// 
			this.toolStrip1.AutoSize = false;
			this.toolStrip1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(234)))), ((int)(((byte)(234)))), ((int)(((byte)(234)))));
			this.toolStrip1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.toolStrip1.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden;
			this.toolStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
			this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.btnCreatePlan,
            this.toolStripSeparator1,
            this.btnEditPlan,
            this.toolStripSeparator7,
            this.btnDelPlan});
			this.toolStrip1.LayoutStyle = System.Windows.Forms.ToolStripLayoutStyle.Flow;
			this.toolStrip1.Location = new System.Drawing.Point(0, 0);
			this.toolStrip1.Name = "toolStrip1";
			this.toolStrip1.Padding = new System.Windows.Forms.Padding(0, 0, 2, 0);
			this.toolStrip1.RenderMode = System.Windows.Forms.ToolStripRenderMode.Professional;
			this.toolStrip1.Size = new System.Drawing.Size(956, 48);
			this.toolStrip1.Stretch = true;
			this.toolStrip1.TabIndex = 27;
			this.toolStrip1.Text = "toolStrip2";
			// 
			// btnCreatePlan
			// 
			this.btnCreatePlan.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnCreatePlan.Image = ((System.Drawing.Image)(resources.GetObject("btnCreatePlan.Image")));
			this.btnCreatePlan.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnCreatePlan.Name = "btnCreatePlan";
			this.btnCreatePlan.Size = new System.Drawing.Size(102, 42);
			this.btnCreatePlan.Tag = "frmProduct_AddProductH";
			this.btnCreatePlan.Text = "Thêm Order";
			this.btnCreatePlan.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnCreatePlan.Click += new System.EventHandler(this.btnCreatePlan_Click);
			// 
			// toolStripSeparator1
			// 
			this.toolStripSeparator1.Name = "toolStripSeparator1";
			this.toolStripSeparator1.Size = new System.Drawing.Size(6, 23);
			// 
			// btnEditPlan
			// 
			this.btnEditPlan.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnEditPlan.Image = ((System.Drawing.Image)(resources.GetObject("btnEditPlan.Image")));
			this.btnEditPlan.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnEditPlan.Name = "btnEditPlan";
			this.btnEditPlan.Size = new System.Drawing.Size(89, 42);
			this.btnEditPlan.Tag = "frmProduct_EditProductH";
			this.btnEditPlan.Text = "Sửa Order";
			this.btnEditPlan.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnEditPlan.Click += new System.EventHandler(this.btnEditPlan_Click);
			// 
			// toolStripSeparator7
			// 
			this.toolStripSeparator7.Name = "toolStripSeparator7";
			this.toolStripSeparator7.Size = new System.Drawing.Size(6, 23);
			// 
			// btnDelPlan
			// 
			this.btnDelPlan.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnDelPlan.Image = ((System.Drawing.Image)(resources.GetObject("btnDelPlan.Image")));
			this.btnDelPlan.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnDelPlan.Name = "btnDelPlan";
			this.btnDelPlan.Size = new System.Drawing.Size(90, 42);
			this.btnDelPlan.Tag = "frmProduct_DeleteProductH";
			this.btnDelPlan.Text = "Xóa Order";
			this.btnDelPlan.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnDelPlan.Click += new System.EventHandler(this.btnDelPlan_Click);
			// 
			// grdData
			// 
			this.grdData.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.grdData.EmbeddedNavigator.Margin = new System.Windows.Forms.Padding(4);
			this.grdData.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.grdData.Location = new System.Drawing.Point(0, 80);
			this.grdData.LookAndFeel.UseDefaultLookAndFeel = false;
			this.grdData.MainView = this.grvData;
			this.grdData.Margin = new System.Windows.Forms.Padding(4);
			this.grdData.Name = "grdData";
			this.grdData.Size = new System.Drawing.Size(956, 370);
			this.grdData.TabIndex = 29;
			this.grdData.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.grvData});
			// 
			// grvData
			// 
			this.grvData.ColumnPanelRowHeight = 60;
			this.grvData.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
            this.colID,
            this.colSalesOrder,
            this.colOrderedQty,
            this.colQtyOfShipOrder,
            this.colCardNo,
            this.colMotorInspSealNo,
            this.colSerialNo,
            this.colArticleID,
            this.colAssemblyOrderNo,
            this.colNoOfJG,
            this.colDescriptions,
            this.colUPR,
            this.colJGDate});
			this.grvData.GridControl = this.grdData;
			this.grvData.HorzScrollVisibility = DevExpress.XtraGrid.Views.Base.ScrollVisibility.Always;
			this.grvData.Name = "grvData";
			this.grvData.OptionsBehavior.AllowAddRows = DevExpress.Utils.DefaultBoolean.False;
			this.grvData.OptionsBehavior.AllowDeleteRows = DevExpress.Utils.DefaultBoolean.False;
			this.grvData.OptionsBehavior.AllowFixedGroups = DevExpress.Utils.DefaultBoolean.False;
			this.grvData.OptionsBehavior.AllowIncrementalSearch = true;
			this.grvData.OptionsBehavior.Editable = false;
			this.grvData.OptionsView.ColumnAutoWidth = false;
			this.grvData.OptionsView.ShowGroupPanel = false;
			this.grvData.RowHeight = 25;
			this.grvData.DoubleClick += new System.EventHandler(this.gvSonPlan_DoubleClick);
			// 
			// colID
			// 
			this.colID.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.colID.AppearanceCell.Options.UseFont = true;
			this.colID.AppearanceCell.Options.UseTextOptions = true;
			this.colID.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colID.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colID.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colID.AppearanceHeader.ForeColor = System.Drawing.Color.Black;
			this.colID.AppearanceHeader.Options.UseFont = true;
			this.colID.AppearanceHeader.Options.UseForeColor = true;
			this.colID.AppearanceHeader.Options.UseTextOptions = true;
			this.colID.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colID.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colID.Caption = "ID";
			this.colID.FieldName = "ID";
			this.colID.MinWidth = 25;
			this.colID.Name = "colID";
			this.colID.OptionsColumn.AllowEdit = false;
			// 
			// colSalesOrder
			// 
			this.colSalesOrder.AppearanceCell.BackColor = System.Drawing.Color.White;
			this.colSalesOrder.AppearanceCell.BackColor2 = System.Drawing.Color.White;
			this.colSalesOrder.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.colSalesOrder.AppearanceCell.Options.UseBackColor = true;
			this.colSalesOrder.AppearanceCell.Options.UseFont = true;
			this.colSalesOrder.AppearanceCell.Options.UseTextOptions = true;
			this.colSalesOrder.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colSalesOrder.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colSalesOrder.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colSalesOrder.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colSalesOrder.AppearanceHeader.ForeColor = System.Drawing.Color.Black;
			this.colSalesOrder.AppearanceHeader.Options.UseFont = true;
			this.colSalesOrder.AppearanceHeader.Options.UseForeColor = true;
			this.colSalesOrder.AppearanceHeader.Options.UseTextOptions = true;
			this.colSalesOrder.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colSalesOrder.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colSalesOrder.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colSalesOrder.Caption = "Sales Order";
			this.colSalesOrder.FieldName = "SalesOrder";
			this.colSalesOrder.MinWidth = 25;
			this.colSalesOrder.Name = "colSalesOrder";
			this.colSalesOrder.OptionsColumn.AllowEdit = false;
			this.colSalesOrder.Visible = true;
			this.colSalesOrder.VisibleIndex = 0;
			this.colSalesOrder.Width = 200;
			// 
			// colOrderedQty
			// 
			this.colOrderedQty.AppearanceCell.BackColor = System.Drawing.Color.White;
			this.colOrderedQty.AppearanceCell.BackColor2 = System.Drawing.Color.White;
			this.colOrderedQty.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.colOrderedQty.AppearanceCell.Options.UseBackColor = true;
			this.colOrderedQty.AppearanceCell.Options.UseFont = true;
			this.colOrderedQty.AppearanceCell.Options.UseTextOptions = true;
			this.colOrderedQty.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colOrderedQty.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colOrderedQty.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colOrderedQty.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colOrderedQty.AppearanceHeader.ForeColor = System.Drawing.Color.Black;
			this.colOrderedQty.AppearanceHeader.Options.UseFont = true;
			this.colOrderedQty.AppearanceHeader.Options.UseForeColor = true;
			this.colOrderedQty.AppearanceHeader.Options.UseTextOptions = true;
			this.colOrderedQty.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colOrderedQty.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colOrderedQty.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colOrderedQty.Caption = "Ordered Qty";
			this.colOrderedQty.FieldName = "OrderedQty";
			this.colOrderedQty.MinWidth = 25;
			this.colOrderedQty.Name = "colOrderedQty";
			this.colOrderedQty.OptionsColumn.AllowEdit = false;
			this.colOrderedQty.Visible = true;
			this.colOrderedQty.VisibleIndex = 1;
			this.colOrderedQty.Width = 92;
			// 
			// colQtyOfShipOrder
			// 
			this.colQtyOfShipOrder.AppearanceCell.BackColor = System.Drawing.Color.White;
			this.colQtyOfShipOrder.AppearanceCell.BackColor2 = System.Drawing.Color.White;
			this.colQtyOfShipOrder.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.colQtyOfShipOrder.AppearanceCell.Options.UseBackColor = true;
			this.colQtyOfShipOrder.AppearanceCell.Options.UseFont = true;
			this.colQtyOfShipOrder.AppearanceCell.Options.UseTextOptions = true;
			this.colQtyOfShipOrder.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colQtyOfShipOrder.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colQtyOfShipOrder.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colQtyOfShipOrder.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colQtyOfShipOrder.AppearanceHeader.ForeColor = System.Drawing.Color.Black;
			this.colQtyOfShipOrder.AppearanceHeader.Options.UseFont = true;
			this.colQtyOfShipOrder.AppearanceHeader.Options.UseForeColor = true;
			this.colQtyOfShipOrder.AppearanceHeader.Options.UseTextOptions = true;
			this.colQtyOfShipOrder.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colQtyOfShipOrder.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colQtyOfShipOrder.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colQtyOfShipOrder.Caption = "Q\'ty Of Ship Order";
			this.colQtyOfShipOrder.FieldName = "QtyOfShipOrder";
			this.colQtyOfShipOrder.MinWidth = 25;
			this.colQtyOfShipOrder.Name = "colQtyOfShipOrder";
			this.colQtyOfShipOrder.OptionsColumn.AllowEdit = false;
			this.colQtyOfShipOrder.Visible = true;
			this.colQtyOfShipOrder.VisibleIndex = 2;
			this.colQtyOfShipOrder.Width = 99;
			// 
			// colCardNo
			// 
			this.colCardNo.AppearanceCell.BackColor = System.Drawing.Color.White;
			this.colCardNo.AppearanceCell.BackColor2 = System.Drawing.Color.White;
			this.colCardNo.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.colCardNo.AppearanceCell.Options.UseBackColor = true;
			this.colCardNo.AppearanceCell.Options.UseFont = true;
			this.colCardNo.AppearanceCell.Options.UseTextOptions = true;
			this.colCardNo.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colCardNo.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colCardNo.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colCardNo.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colCardNo.AppearanceHeader.ForeColor = System.Drawing.Color.Black;
			this.colCardNo.AppearanceHeader.Options.UseFont = true;
			this.colCardNo.AppearanceHeader.Options.UseForeColor = true;
			this.colCardNo.AppearanceHeader.Options.UseTextOptions = true;
			this.colCardNo.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colCardNo.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colCardNo.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colCardNo.Caption = "Card No";
			this.colCardNo.FieldName = "CardNo";
			this.colCardNo.MinWidth = 25;
			this.colCardNo.Name = "colCardNo";
			this.colCardNo.OptionsColumn.AllowEdit = false;
			this.colCardNo.Visible = true;
			this.colCardNo.VisibleIndex = 3;
			this.colCardNo.Width = 143;
			// 
			// colMotorInspSealNo
			// 
			this.colMotorInspSealNo.AppearanceCell.BackColor = System.Drawing.Color.White;
			this.colMotorInspSealNo.AppearanceCell.BackColor2 = System.Drawing.Color.White;
			this.colMotorInspSealNo.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.colMotorInspSealNo.AppearanceCell.Options.UseBackColor = true;
			this.colMotorInspSealNo.AppearanceCell.Options.UseFont = true;
			this.colMotorInspSealNo.AppearanceCell.Options.UseTextOptions = true;
			this.colMotorInspSealNo.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colMotorInspSealNo.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colMotorInspSealNo.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colMotorInspSealNo.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colMotorInspSealNo.AppearanceHeader.ForeColor = System.Drawing.Color.Black;
			this.colMotorInspSealNo.AppearanceHeader.Options.UseFont = true;
			this.colMotorInspSealNo.AppearanceHeader.Options.UseForeColor = true;
			this.colMotorInspSealNo.AppearanceHeader.Options.UseTextOptions = true;
			this.colMotorInspSealNo.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colMotorInspSealNo.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colMotorInspSealNo.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colMotorInspSealNo.Caption = "Motor_Insp_Seal No";
			this.colMotorInspSealNo.FieldName = "MotorInspSealNo";
			this.colMotorInspSealNo.MinWidth = 25;
			this.colMotorInspSealNo.Name = "colMotorInspSealNo";
			this.colMotorInspSealNo.OptionsColumn.AllowEdit = false;
			this.colMotorInspSealNo.Visible = true;
			this.colMotorInspSealNo.VisibleIndex = 4;
			this.colMotorInspSealNo.Width = 154;
			// 
			// colSerialNo
			// 
			this.colSerialNo.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.colSerialNo.AppearanceCell.Options.UseFont = true;
			this.colSerialNo.AppearanceCell.Options.UseTextOptions = true;
			this.colSerialNo.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colSerialNo.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colSerialNo.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colSerialNo.AppearanceHeader.ForeColor = System.Drawing.Color.Black;
			this.colSerialNo.AppearanceHeader.Options.UseFont = true;
			this.colSerialNo.AppearanceHeader.Options.UseForeColor = true;
			this.colSerialNo.AppearanceHeader.Options.UseTextOptions = true;
			this.colSerialNo.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colSerialNo.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colSerialNo.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colSerialNo.Caption = "Serial No";
			this.colSerialNo.FieldName = "SerialNo";
			this.colSerialNo.MinWidth = 25;
			this.colSerialNo.Name = "colSerialNo";
			this.colSerialNo.Visible = true;
			this.colSerialNo.VisibleIndex = 5;
			// 
			// colArticleID
			// 
			this.colArticleID.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.colArticleID.AppearanceCell.Options.UseFont = true;
			this.colArticleID.AppearanceCell.Options.UseTextOptions = true;
			this.colArticleID.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colArticleID.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colArticleID.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colArticleID.AppearanceHeader.ForeColor = System.Drawing.Color.Black;
			this.colArticleID.AppearanceHeader.Options.UseFont = true;
			this.colArticleID.AppearanceHeader.Options.UseForeColor = true;
			this.colArticleID.AppearanceHeader.Options.UseTextOptions = true;
			this.colArticleID.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colArticleID.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colArticleID.Caption = "Article ID";
			this.colArticleID.FieldName = "ArticleID";
			this.colArticleID.MinWidth = 25;
			this.colArticleID.Name = "colArticleID";
			this.colArticleID.Visible = true;
			this.colArticleID.VisibleIndex = 6;
			this.colArticleID.Width = 196;
			// 
			// colAssemblyOrderNo
			// 
			this.colAssemblyOrderNo.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.colAssemblyOrderNo.AppearanceCell.Options.UseFont = true;
			this.colAssemblyOrderNo.AppearanceCell.Options.UseTextOptions = true;
			this.colAssemblyOrderNo.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colAssemblyOrderNo.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colAssemblyOrderNo.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colAssemblyOrderNo.AppearanceHeader.ForeColor = System.Drawing.Color.Black;
			this.colAssemblyOrderNo.AppearanceHeader.Options.UseFont = true;
			this.colAssemblyOrderNo.AppearanceHeader.Options.UseForeColor = true;
			this.colAssemblyOrderNo.AppearanceHeader.Options.UseTextOptions = true;
			this.colAssemblyOrderNo.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colAssemblyOrderNo.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colAssemblyOrderNo.Caption = "Assembly Order No";
			this.colAssemblyOrderNo.FieldName = "AssemblyOrderNo";
			this.colAssemblyOrderNo.MinWidth = 25;
			this.colAssemblyOrderNo.Name = "colAssemblyOrderNo";
			this.colAssemblyOrderNo.Visible = true;
			this.colAssemblyOrderNo.VisibleIndex = 7;
			this.colAssemblyOrderNo.Width = 173;
			// 
			// colNoOfJG
			// 
			this.colNoOfJG.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.colNoOfJG.AppearanceCell.Options.UseFont = true;
			this.colNoOfJG.AppearanceCell.Options.UseTextOptions = true;
			this.colNoOfJG.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colNoOfJG.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colNoOfJG.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colNoOfJG.AppearanceHeader.ForeColor = System.Drawing.Color.Black;
			this.colNoOfJG.AppearanceHeader.Options.UseFont = true;
			this.colNoOfJG.AppearanceHeader.Options.UseForeColor = true;
			this.colNoOfJG.AppearanceHeader.Options.UseTextOptions = true;
			this.colNoOfJG.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colNoOfJG.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colNoOfJG.Caption = "No Of J/G";
			this.colNoOfJG.FieldName = "NoOfJG";
			this.colNoOfJG.MinWidth = 25;
			this.colNoOfJG.Name = "colNoOfJG";
			this.colNoOfJG.Visible = true;
			this.colNoOfJG.VisibleIndex = 8;
			this.colNoOfJG.Width = 161;
			// 
			// colDescriptions
			// 
			this.colDescriptions.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.colDescriptions.AppearanceCell.Options.UseFont = true;
			this.colDescriptions.AppearanceCell.Options.UseTextOptions = true;
			this.colDescriptions.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colDescriptions.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colDescriptions.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colDescriptions.AppearanceHeader.ForeColor = System.Drawing.Color.Black;
			this.colDescriptions.AppearanceHeader.Options.UseFont = true;
			this.colDescriptions.AppearanceHeader.Options.UseForeColor = true;
			this.colDescriptions.AppearanceHeader.Options.UseTextOptions = true;
			this.colDescriptions.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colDescriptions.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colDescriptions.Caption = "Descriptions";
			this.colDescriptions.FieldName = "Descriptions";
			this.colDescriptions.MinWidth = 25;
			this.colDescriptions.Name = "colDescriptions";
			this.colDescriptions.Visible = true;
			this.colDescriptions.VisibleIndex = 9;
			// 
			// colUPR
			// 
			this.colUPR.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.colUPR.AppearanceCell.Options.UseFont = true;
			this.colUPR.AppearanceCell.Options.UseTextOptions = true;
			this.colUPR.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colUPR.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colUPR.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colUPR.AppearanceHeader.ForeColor = System.Drawing.Color.Black;
			this.colUPR.AppearanceHeader.Options.UseFont = true;
			this.colUPR.AppearanceHeader.Options.UseForeColor = true;
			this.colUPR.AppearanceHeader.Options.UseTextOptions = true;
			this.colUPR.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colUPR.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colUPR.Caption = "U/P/R";
			this.colUPR.FieldName = "UPR";
			this.colUPR.MinWidth = 25;
			this.colUPR.Name = "colUPR";
			this.colUPR.Visible = true;
			this.colUPR.VisibleIndex = 10;
			this.colUPR.Width = 118;
			// 
			// colJGDate
			// 
			this.colJGDate.AppearanceCell.Font = new System.Drawing.Font("Tahoma", 11F);
			this.colJGDate.AppearanceCell.Options.UseFont = true;
			this.colJGDate.AppearanceCell.Options.UseTextOptions = true;
			this.colJGDate.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colJGDate.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colJGDate.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colJGDate.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.colJGDate.AppearanceHeader.Options.UseFont = true;
			this.colJGDate.AppearanceHeader.Options.UseTextOptions = true;
			this.colJGDate.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colJGDate.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colJGDate.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colJGDate.Caption = "JG Date";
			this.colJGDate.DisplayFormat.FormatString = "dd/MM/yyyy";
			this.colJGDate.DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
			this.colJGDate.FieldName = "JGDate";
			this.colJGDate.Name = "colJGDate";
			this.colJGDate.Visible = true;
			this.colJGDate.VisibleIndex = 11;
			this.colJGDate.Width = 124;
			// 
			// panel2
			// 
			this.panel2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.panel2.Controls.Add(this.btnSearch);
			this.panel2.Controls.Add(this.txbSearch);
			this.panel2.Controls.Add(this.dtpFrom);
			this.panel2.Controls.Add(this.dtpTo);
			this.panel2.Controls.Add(this.label1);
			this.panel2.Controls.Add(this.label3);
			this.panel2.Location = new System.Drawing.Point(0, 43);
			this.panel2.Name = "panel2";
			this.panel2.Size = new System.Drawing.Size(630, 34);
			this.panel2.TabIndex = 30;
			// 
			// btnSearch
			// 
			this.btnSearch.Appearance.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnSearch.Appearance.Options.UseFont = true;
			this.btnSearch.Location = new System.Drawing.Point(547, 5);
			this.btnSearch.Name = "btnSearch";
			this.btnSearch.Size = new System.Drawing.Size(75, 23);
			this.btnSearch.TabIndex = 25;
			this.btnSearch.Text = "Tìm kiếm";
			this.btnSearch.Click += new System.EventHandler(this.btnSearch_Click);
			// 
			// txbSearch
			// 
			this.txbSearch.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txbSearch.Location = new System.Drawing.Point(394, 6);
			this.txbSearch.Name = "txbSearch";
			this.txbSearch.Size = new System.Drawing.Size(147, 24);
			this.txbSearch.TabIndex = 24;
			this.txbSearch.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txbSearch_KeyDown);
			// 
			// dtpFrom
			// 
			this.dtpFrom.CustomFormat = "dd/MM/yyyy";
			this.dtpFrom.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.dtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
			this.dtpFrom.Location = new System.Drawing.Point(77, 6);
			this.dtpFrom.Name = "dtpFrom";
			this.dtpFrom.Size = new System.Drawing.Size(106, 24);
			this.dtpFrom.TabIndex = 36;
			this.dtpFrom.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtpFrom_KeyPress);
			// 
			// dtpTo
			// 
			this.dtpTo.CustomFormat = "dd/MM/yyyy";
			this.dtpTo.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.dtpTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
			this.dtpTo.Location = new System.Drawing.Point(274, 6);
			this.dtpTo.Name = "dtpTo";
			this.dtpTo.Size = new System.Drawing.Size(99, 24);
			this.dtpTo.TabIndex = 35;
			this.dtpTo.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtpTo_KeyPress);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.label1.Location = new System.Drawing.Point(205, 8);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(78, 18);
			this.label1.TabIndex = 34;
			this.label1.Text = "Đến ngày: ";
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.label3.Location = new System.Drawing.Point(12, 8);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(68, 18);
			this.label3.TabIndex = 33;
			this.label3.Text = "Từ ngày: ";
			// 
			// frmCheckMotor
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(956, 450);
			this.Controls.Add(this.panel2);
			this.Controls.Add(this.grdData);
			this.Controls.Add(this.toolStrip1);
			this.Margin = new System.Windows.Forms.Padding(4);
			this.Name = "frmCheckMotor";
			this.Text = "DANH SÁCH LẮP RÁP THEO ORDER";
			this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
			this.Load += new System.EventHandler(this.frmCheckMotor_Load);
			this.toolStrip1.ResumeLayout(false);
			this.toolStrip1.PerformLayout();
			((System.ComponentModel.ISupportInitialize)(this.grdData)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grvData)).EndInit();
			this.panel2.ResumeLayout(false);
			this.panel2.PerformLayout();
			this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ToolStrip toolStrip1;
        private System.Windows.Forms.ToolStripButton btnCreatePlan;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripButton btnEditPlan;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator7;
        private System.Windows.Forms.ToolStripButton btnDelPlan;
        private DevExpress.XtraGrid.GridControl grdData;
        private DevExpress.XtraGrid.Views.Grid.GridView grvData;
        private DevExpress.XtraGrid.Columns.GridColumn colID;
        private DevExpress.XtraGrid.Columns.GridColumn colSalesOrder;
        private DevExpress.XtraGrid.Columns.GridColumn colOrderedQty;
        private DevExpress.XtraGrid.Columns.GridColumn colQtyOfShipOrder;
        private DevExpress.XtraGrid.Columns.GridColumn colCardNo;
        private DevExpress.XtraGrid.Columns.GridColumn colMotorInspSealNo;
		private DevExpress.XtraGrid.Columns.GridColumn colSerialNo;
		private DevExpress.XtraGrid.Columns.GridColumn colArticleID;
		private DevExpress.XtraGrid.Columns.GridColumn colAssemblyOrderNo;
		private DevExpress.XtraGrid.Columns.GridColumn colNoOfJG;
		private DevExpress.XtraGrid.Columns.GridColumn colDescriptions;
		private DevExpress.XtraGrid.Columns.GridColumn colUPR;
		private System.Windows.Forms.Timer timer1;
		private System.Windows.Forms.Panel panel2;
		private DevExpress.XtraEditors.SimpleButton btnSearch;
		private System.Windows.Forms.TextBox txbSearch;
		private System.Windows.Forms.DateTimePicker dtpFrom;
		private System.Windows.Forms.DateTimePicker dtpTo;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label3;
		private DevExpress.XtraGrid.Columns.GridColumn colJGDate;
	}
}