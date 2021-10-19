
namespace BMS
{
	partial class frmCheckMotorDetails
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
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmCheckMotorDetails));
			this.mnuMenu = new System.Windows.Forms.ToolStrip();
			this.btnSaveClose = new System.Windows.Forms.ToolStripButton();
			this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
			this.btnSaveNew = new System.Windows.Forms.ToolStripButton();
			this.menuStrip1 = new System.Windows.Forms.MenuStrip();
			this.cấtToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.catVaThemOiToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.labelControl3 = new DevExpress.XtraEditors.LabelControl();
			this.txtAssemblyOrderNo = new System.Windows.Forms.TextBox();
			this.labelControl16 = new DevExpress.XtraEditors.LabelControl();
			this.labelControl4 = new DevExpress.XtraEditors.LabelControl();
			this.txtSalesOrder = new System.Windows.Forms.TextBox();
			this.labelControl9 = new DevExpress.XtraEditors.LabelControl();
			this.txtArticleID = new System.Windows.Forms.TextBox();
			this.txtCardNo = new System.Windows.Forms.TextBox();
			this.labelControl14 = new DevExpress.XtraEditors.LabelControl();
			this.labelControl1 = new DevExpress.XtraEditors.LabelControl();
			this.labelControl2 = new DevExpress.XtraEditors.LabelControl();
			this.labelControl10 = new DevExpress.XtraEditors.LabelControl();
			this.txtMotorInspSealNo = new System.Windows.Forms.TextBox();
			this.txtDescription = new System.Windows.Forms.RichTextBox();
			this.txtQtyOfShipOrder = new System.Windows.Forms.TextBox();
			this.txtSerialNo = new System.Windows.Forms.TextBox();
			this.txtOrderedQty = new System.Windows.Forms.TextBox();
			this.lbLot = new DevExpress.XtraEditors.LabelControl();
			this.txtNoOfJG = new System.Windows.Forms.TextBox();
			this.labelControl5 = new DevExpress.XtraEditors.LabelControl();
			this.labelControl6 = new DevExpress.XtraEditors.LabelControl();
			this.txtUPR = new System.Windows.Forms.TextBox();
			this.mnuMenu.SuspendLayout();
			this.menuStrip1.SuspendLayout();
			this.SuspendLayout();
			// 
			// mnuMenu
			// 
			this.mnuMenu.AutoSize = false;
			this.mnuMenu.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(234)))), ((int)(((byte)(234)))), ((int)(((byte)(234)))));
			this.mnuMenu.Font = new System.Drawing.Font("Tahoma", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.mnuMenu.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden;
			this.mnuMenu.ImageScalingSize = new System.Drawing.Size(20, 20);
			this.mnuMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.btnSaveClose,
            this.toolStripSeparator2,
            this.btnSaveNew});
			this.mnuMenu.LayoutStyle = System.Windows.Forms.ToolStripLayoutStyle.Flow;
			this.mnuMenu.Location = new System.Drawing.Point(0, 0);
			this.mnuMenu.Name = "mnuMenu";
			this.mnuMenu.RenderMode = System.Windows.Forms.ToolStripRenderMode.Professional;
			this.mnuMenu.Size = new System.Drawing.Size(1175, 46);
			this.mnuMenu.TabIndex = 2;
			this.mnuMenu.Text = "toolStrip2";
			// 
			// btnSaveClose
			// 
			this.btnSaveClose.Font = new System.Drawing.Font("Tahoma", 10.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnSaveClose.Image = ((System.Drawing.Image)(resources.GetObject("btnSaveClose.Image")));
			this.btnSaveClose.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnSaveClose.Name = "btnSaveClose";
			this.btnSaveClose.Size = new System.Drawing.Size(96, 42);
			this.btnSaveClose.Tag = "frmProduct_AddProductH";
			this.btnSaveClose.Text = "Cất && Đóng";
			this.btnSaveClose.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnSaveClose.Click += new System.EventHandler(this.btnSaveClose_Click);
			// 
			// toolStripSeparator2
			// 
			this.toolStripSeparator2.AutoSize = false;
			this.toolStripSeparator2.Name = "toolStripSeparator2";
			this.toolStripSeparator2.Size = new System.Drawing.Size(6, 40);
			// 
			// btnSaveNew
			// 
			this.btnSaveNew.Font = new System.Drawing.Font("Tahoma", 10.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnSaveNew.Image = ((System.Drawing.Image)(resources.GetObject("btnSaveNew.Image")));
			this.btnSaveNew.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnSaveNew.Name = "btnSaveNew";
			this.btnSaveNew.Size = new System.Drawing.Size(128, 42);
			this.btnSaveNew.Tag = "frmProduct_AddProductH";
			this.btnSaveNew.Text = "Cất && Thêm mới";
			this.btnSaveNew.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnSaveNew.Click += new System.EventHandler(this.btnSaveNew_Click);
			// 
			// menuStrip1
			// 
			this.menuStrip1.Dock = System.Windows.Forms.DockStyle.None;
			this.menuStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
			this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.cấtToolStripMenuItem,
            this.catVaThemOiToolStripMenuItem});
			this.menuStrip1.Location = new System.Drawing.Point(430, 7);
			this.menuStrip1.Name = "menuStrip1";
			this.menuStrip1.Padding = new System.Windows.Forms.Padding(4, 2, 0, 2);
			this.menuStrip1.Size = new System.Drawing.Size(139, 24);
			this.menuStrip1.TabIndex = 225;
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
			// catVaThemOiToolStripMenuItem
			// 
			this.catVaThemOiToolStripMenuItem.Name = "catVaThemOiToolStripMenuItem";
			this.catVaThemOiToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.S)));
			this.catVaThemOiToolStripMenuItem.Size = new System.Drawing.Size(96, 20);
			this.catVaThemOiToolStripMenuItem.Text = "Cat va them oi";
			this.catVaThemOiToolStripMenuItem.Click += new System.EventHandler(this.catVaThemOiToolStripMenuItem_Click);
			// 
			// labelControl3
			// 
			this.labelControl3.Appearance.Font = new System.Drawing.Font("Microsoft Sans Serif", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.labelControl3.Location = new System.Drawing.Point(11, 56);
			this.labelControl3.Name = "labelControl3";
			this.labelControl3.Size = new System.Drawing.Size(90, 29);
			this.labelControl3.TabIndex = 229;
			this.labelControl3.Text = "Card No";
			// 
			// txtAssemblyOrderNo
			// 
			this.txtAssemblyOrderNo.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.txtAssemblyOrderNo.Font = new System.Drawing.Font("Microsoft Sans Serif", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtAssemblyOrderNo.Location = new System.Drawing.Point(835, 212);
			this.txtAssemblyOrderNo.Name = "txtAssemblyOrderNo";
			this.txtAssemblyOrderNo.Size = new System.Drawing.Size(322, 35);
			this.txtAssemblyOrderNo.TabIndex = 240;
			// 
			// labelControl16
			// 
			this.labelControl16.Appearance.Font = new System.Drawing.Font("Microsoft Sans Serif", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.labelControl16.Location = new System.Drawing.Point(602, 215);
			this.labelControl16.Name = "labelControl16";
			this.labelControl16.Size = new System.Drawing.Size(211, 29);
			this.labelControl16.TabIndex = 241;
			this.labelControl16.Text = "Assembly Order No";
			// 
			// labelControl4
			// 
			this.labelControl4.Appearance.Font = new System.Drawing.Font("Microsoft Sans Serif", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.labelControl4.Location = new System.Drawing.Point(11, 160);
			this.labelControl4.Name = "labelControl4";
			this.labelControl4.Size = new System.Drawing.Size(130, 29);
			this.labelControl4.TabIndex = 230;
			this.labelControl4.Text = "Sales Order";
			// 
			// txtSalesOrder
			// 
			this.txtSalesOrder.Font = new System.Drawing.Font("Microsoft Sans Serif", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtSalesOrder.Location = new System.Drawing.Point(160, 157);
			this.txtSalesOrder.Name = "txtSalesOrder";
			this.txtSalesOrder.Size = new System.Drawing.Size(411, 35);
			this.txtSalesOrder.TabIndex = 228;
			// 
			// labelControl9
			// 
			this.labelControl9.Appearance.Font = new System.Drawing.Font("Microsoft Sans Serif", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.labelControl9.Location = new System.Drawing.Point(11, 108);
			this.labelControl9.Name = "labelControl9";
			this.labelControl9.Size = new System.Drawing.Size(96, 29);
			this.labelControl9.TabIndex = 231;
			this.labelControl9.Text = "Article ID";
			// 
			// txtArticleID
			// 
			this.txtArticleID.Font = new System.Drawing.Font("Microsoft Sans Serif", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtArticleID.Location = new System.Drawing.Point(160, 105);
			this.txtArticleID.Name = "txtArticleID";
			this.txtArticleID.Size = new System.Drawing.Size(411, 35);
			this.txtArticleID.TabIndex = 227;
			// 
			// txtCardNo
			// 
			this.txtCardNo.Font = new System.Drawing.Font("Microsoft Sans Serif", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtCardNo.Location = new System.Drawing.Point(160, 53);
			this.txtCardNo.Name = "txtCardNo";
			this.txtCardNo.Size = new System.Drawing.Size(411, 35);
			this.txtCardNo.TabIndex = 226;
			// 
			// labelControl14
			// 
			this.labelControl14.Appearance.Font = new System.Drawing.Font("Microsoft Sans Serif", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.labelControl14.Location = new System.Drawing.Point(602, 161);
			this.labelControl14.Name = "labelControl14";
			this.labelControl14.Size = new System.Drawing.Size(101, 29);
			this.labelControl14.TabIndex = 238;
			this.labelControl14.Text = "Serial No";
			// 
			// labelControl1
			// 
			this.labelControl1.Appearance.Font = new System.Drawing.Font("Microsoft Sans Serif", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.labelControl1.Location = new System.Drawing.Point(11, 217);
			this.labelControl1.Name = "labelControl1";
			this.labelControl1.Size = new System.Drawing.Size(122, 29);
			this.labelControl1.TabIndex = 232;
			this.labelControl1.Text = "Description";
			// 
			// labelControl2
			// 
			this.labelControl2.Appearance.Font = new System.Drawing.Font("Microsoft Sans Serif", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.labelControl2.Location = new System.Drawing.Point(602, 56);
			this.labelControl2.Name = "labelControl2";
			this.labelControl2.Size = new System.Drawing.Size(195, 29);
			this.labelControl2.TabIndex = 234;
			this.labelControl2.Text = "Q\'ty Of Ship Order";
			// 
			// labelControl10
			// 
			this.labelControl10.Appearance.Font = new System.Drawing.Font("Microsoft Sans Serif", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.labelControl10.Location = new System.Drawing.Point(602, 109);
			this.labelControl10.Name = "labelControl10";
			this.labelControl10.Size = new System.Drawing.Size(220, 29);
			this.labelControl10.TabIndex = 237;
			this.labelControl10.Text = "Motor_Insp_Seal No";
			// 
			// txtMotorInspSealNo
			// 
			this.txtMotorInspSealNo.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.txtMotorInspSealNo.Font = new System.Drawing.Font("Microsoft Sans Serif", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtMotorInspSealNo.Location = new System.Drawing.Point(835, 106);
			this.txtMotorInspSealNo.Name = "txtMotorInspSealNo";
			this.txtMotorInspSealNo.Size = new System.Drawing.Size(322, 35);
			this.txtMotorInspSealNo.TabIndex = 236;
			// 
			// txtDescription
			// 
			this.txtDescription.Location = new System.Drawing.Point(160, 215);
			this.txtDescription.Name = "txtDescription";
			this.txtDescription.Size = new System.Drawing.Size(411, 35);
			this.txtDescription.TabIndex = 242;
			this.txtDescription.Text = "";
			// 
			// txtQtyOfShipOrder
			// 
			this.txtQtyOfShipOrder.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.txtQtyOfShipOrder.Font = new System.Drawing.Font("Microsoft Sans Serif", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtQtyOfShipOrder.Location = new System.Drawing.Point(835, 53);
			this.txtQtyOfShipOrder.Name = "txtQtyOfShipOrder";
			this.txtQtyOfShipOrder.Size = new System.Drawing.Size(322, 35);
			this.txtQtyOfShipOrder.TabIndex = 243;
			// 
			// txtSerialNo
			// 
			this.txtSerialNo.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.txtSerialNo.Font = new System.Drawing.Font("Microsoft Sans Serif", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtSerialNo.Location = new System.Drawing.Point(835, 158);
			this.txtSerialNo.Name = "txtSerialNo";
			this.txtSerialNo.Size = new System.Drawing.Size(322, 35);
			this.txtSerialNo.TabIndex = 244;
			// 
			// txtOrderedQty
			// 
			this.txtOrderedQty.Font = new System.Drawing.Font("Microsoft Sans Serif", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtOrderedQty.Location = new System.Drawing.Point(160, 267);
			this.txtOrderedQty.Name = "txtOrderedQty";
			this.txtOrderedQty.Size = new System.Drawing.Size(411, 35);
			this.txtOrderedQty.TabIndex = 246;
			// 
			// lbLot
			// 
			this.lbLot.Appearance.Font = new System.Drawing.Font("Microsoft Sans Serif", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lbLot.Location = new System.Drawing.Point(11, 270);
			this.lbLot.Name = "lbLot";
			this.lbLot.Size = new System.Drawing.Size(138, 29);
			this.lbLot.TabIndex = 245;
			this.lbLot.Text = "Ordered Q\'ty";
			// 
			// txtNoOfJG
			// 
			this.txtNoOfJG.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.txtNoOfJG.Font = new System.Drawing.Font("Microsoft Sans Serif", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtNoOfJG.Location = new System.Drawing.Point(835, 264);
			this.txtNoOfJG.Name = "txtNoOfJG";
			this.txtNoOfJG.Size = new System.Drawing.Size(322, 35);
			this.txtNoOfJG.TabIndex = 236;
			// 
			// labelControl5
			// 
			this.labelControl5.Appearance.Font = new System.Drawing.Font("Microsoft Sans Serif", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.labelControl5.Location = new System.Drawing.Point(602, 267);
			this.labelControl5.Name = "labelControl5";
			this.labelControl5.Size = new System.Drawing.Size(106, 29);
			this.labelControl5.TabIndex = 237;
			this.labelControl5.Text = "No Of J/G";
			// 
			// labelControl6
			// 
			this.labelControl6.Appearance.Font = new System.Drawing.Font("Microsoft Sans Serif", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.labelControl6.Location = new System.Drawing.Point(602, 319);
			this.labelControl6.Name = "labelControl6";
			this.labelControl6.Size = new System.Drawing.Size(64, 29);
			this.labelControl6.TabIndex = 238;
			this.labelControl6.Text = "U/P/R";
			// 
			// txtUPR
			// 
			this.txtUPR.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.txtUPR.Font = new System.Drawing.Font("Microsoft Sans Serif", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtUPR.Location = new System.Drawing.Point(835, 316);
			this.txtUPR.Name = "txtUPR";
			this.txtUPR.Size = new System.Drawing.Size(322, 35);
			this.txtUPR.TabIndex = 244;
			// 
			// frmCheckMotorDetails
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(1175, 364);
			this.Controls.Add(this.txtOrderedQty);
			this.Controls.Add(this.lbLot);
			this.Controls.Add(this.txtUPR);
			this.Controls.Add(this.txtSerialNo);
			this.Controls.Add(this.txtQtyOfShipOrder);
			this.Controls.Add(this.txtDescription);
			this.Controls.Add(this.labelControl3);
			this.Controls.Add(this.txtAssemblyOrderNo);
			this.Controls.Add(this.labelControl16);
			this.Controls.Add(this.labelControl4);
			this.Controls.Add(this.txtSalesOrder);
			this.Controls.Add(this.labelControl9);
			this.Controls.Add(this.txtArticleID);
			this.Controls.Add(this.labelControl6);
			this.Controls.Add(this.txtCardNo);
			this.Controls.Add(this.labelControl14);
			this.Controls.Add(this.labelControl1);
			this.Controls.Add(this.labelControl5);
			this.Controls.Add(this.labelControl2);
			this.Controls.Add(this.txtNoOfJG);
			this.Controls.Add(this.labelControl10);
			this.Controls.Add(this.txtMotorInspSealNo);
			this.Controls.Add(this.menuStrip1);
			this.Controls.Add(this.mnuMenu);
			this.Margin = new System.Windows.Forms.Padding(2);
			this.Name = "frmCheckMotorDetails";
			this.Text = "frmAddEditPlan";
			this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmAddEditPlan_FormClosing);
			this.Load += new System.EventHandler(this.frmAddEditPlan_Load);
			this.mnuMenu.ResumeLayout(false);
			this.mnuMenu.PerformLayout();
			this.menuStrip1.ResumeLayout(false);
			this.menuStrip1.PerformLayout();
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.ToolStrip mnuMenu;
		private System.Windows.Forms.ToolStripButton btnSaveClose;
		private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
		private System.Windows.Forms.ToolStripButton btnSaveNew;
		private System.Windows.Forms.MenuStrip menuStrip1;
		private System.Windows.Forms.ToolStripMenuItem cấtToolStripMenuItem;
		private System.Windows.Forms.ToolStripMenuItem catVaThemOiToolStripMenuItem;
		private DevExpress.XtraEditors.LabelControl labelControl3;
		private System.Windows.Forms.TextBox txtAssemblyOrderNo;
		private DevExpress.XtraEditors.LabelControl labelControl16;
		private DevExpress.XtraEditors.LabelControl labelControl4;
		private System.Windows.Forms.TextBox txtSalesOrder;
		private DevExpress.XtraEditors.LabelControl labelControl9;
		private System.Windows.Forms.TextBox txtArticleID;
		private System.Windows.Forms.TextBox txtCardNo;
		private DevExpress.XtraEditors.LabelControl labelControl14;
		private DevExpress.XtraEditors.LabelControl labelControl1;
		private DevExpress.XtraEditors.LabelControl labelControl2;
		private DevExpress.XtraEditors.LabelControl labelControl10;
		private System.Windows.Forms.TextBox txtMotorInspSealNo;
		private System.Windows.Forms.RichTextBox txtDescription;
		private System.Windows.Forms.TextBox txtQtyOfShipOrder;
		private System.Windows.Forms.TextBox txtSerialNo;
		private System.Windows.Forms.TextBox txtOrderedQty;
		private DevExpress.XtraEditors.LabelControl lbLot;
		private System.Windows.Forms.TextBox txtNoOfJG;
		private DevExpress.XtraEditors.LabelControl labelControl5;
		private DevExpress.XtraEditors.LabelControl labelControl6;
		private System.Windows.Forms.TextBox txtUPR;
	}
}