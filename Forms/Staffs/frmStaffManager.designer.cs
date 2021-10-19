namespace BMS
{
    partial class frmStaffManager
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
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmStaffManager));
			this.mnuMenu = new System.Windows.Forms.ToolStrip();
			this.btnNew = new System.Windows.Forms.ToolStripButton();
			this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
			this.btnEdit = new System.Windows.Forms.ToolStripButton();
			this.toolStripSeparator3 = new System.Windows.Forms.ToolStripSeparator();
			this.btnDelete = new System.Windows.Forms.ToolStripButton();
			this.toolStripSeparator4 = new System.Windows.Forms.ToolStripSeparator();
			this.toolStripSeparator7 = new System.Windows.Forms.ToolStripSeparator();
			this.btnExcel = new System.Windows.Forms.ToolStripButton();
			this.toolStripSeparator5 = new System.Windows.Forms.ToolStripSeparator();
			this.btnQLNhom = new System.Windows.Forms.ToolStripButton();
			this.toolStripSeparator8 = new System.Windows.Forms.ToolStripSeparator();
			this.grdData = new DevExpress.XtraGrid.GridControl();
			this.grvData = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.colID = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colCode = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colFullName = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colBirthOfDate = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colHomeAddress = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colQualifications = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colHandPhone = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colMST = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colBHYT = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colBankAccount = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colSex = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colDepartmentID = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colDepartmentName = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colGroup = new DevExpress.XtraGrid.Columns.GridColumn();
			this.NhomPhong = new DevExpress.XtraEditors.Repository.RepositoryItemLookUpEdit();
			this.colLoginName = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colStatus = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colEmail = new DevExpress.XtraGrid.Columns.GridColumn();
			this.label2 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.mnuMenu.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.grdData)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grvData)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.NhomPhong)).BeginInit();
			this.SuspendLayout();
			// 
			// mnuMenu
			// 
			this.mnuMenu.AutoSize = false;
			this.mnuMenu.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(234)))), ((int)(((byte)(234)))), ((int)(((byte)(234)))));
			this.mnuMenu.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.mnuMenu.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden;
			this.mnuMenu.ImageScalingSize = new System.Drawing.Size(20, 20);
			this.mnuMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.btnNew,
            this.toolStripSeparator2,
            this.btnEdit,
            this.toolStripSeparator3,
            this.btnDelete,
            this.toolStripSeparator4,
            this.toolStripSeparator7,
            this.btnExcel,
            this.toolStripSeparator5,
            this.btnQLNhom,
            this.toolStripSeparator8});
			this.mnuMenu.LayoutStyle = System.Windows.Forms.ToolStripLayoutStyle.Flow;
			this.mnuMenu.Location = new System.Drawing.Point(0, 0);
			this.mnuMenu.Name = "mnuMenu";
			this.mnuMenu.RenderMode = System.Windows.Forms.ToolStripRenderMode.Professional;
			this.mnuMenu.Size = new System.Drawing.Size(1248, 42);
			this.mnuMenu.TabIndex = 14;
			this.mnuMenu.Text = "toolStrip2";
			// 
			// btnNew
			// 
			this.btnNew.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnNew.Image = ((System.Drawing.Image)(resources.GetObject("btnNew.Image")));
			this.btnNew.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnNew.Name = "btnNew";
			this.btnNew.Size = new System.Drawing.Size(43, 37);
			this.btnNew.Tag = "Staff_Add";
			this.btnNew.Text = "Thêm";
			this.btnNew.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnNew.Click += new System.EventHandler(this.btnNew_Click);
			// 
			// toolStripSeparator2
			// 
			this.toolStripSeparator2.Name = "toolStripSeparator2";
			this.toolStripSeparator2.Size = new System.Drawing.Size(6, 23);
			// 
			// btnEdit
			// 
			this.btnEdit.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnEdit.Image = ((System.Drawing.Image)(resources.GetObject("btnEdit.Image")));
			this.btnEdit.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnEdit.Name = "btnEdit";
			this.btnEdit.Size = new System.Drawing.Size(33, 37);
			this.btnEdit.Tag = "Staff_Edit";
			this.btnEdit.Text = "Sửa";
			this.btnEdit.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnEdit.Click += new System.EventHandler(this.btnEdit_Click);
			// 
			// toolStripSeparator3
			// 
			this.toolStripSeparator3.Name = "toolStripSeparator3";
			this.toolStripSeparator3.Size = new System.Drawing.Size(6, 23);
			// 
			// btnDelete
			// 
			this.btnDelete.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnDelete.Image = ((System.Drawing.Image)(resources.GetObject("btnDelete.Image")));
			this.btnDelete.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnDelete.Name = "btnDelete";
			this.btnDelete.Size = new System.Drawing.Size(32, 37);
			this.btnDelete.Tag = "Staff_Del";
			this.btnDelete.Text = "Xóa";
			this.btnDelete.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnDelete.Click += new System.EventHandler(this.btnDelete_Click);
			// 
			// toolStripSeparator4
			// 
			this.toolStripSeparator4.Name = "toolStripSeparator4";
			this.toolStripSeparator4.Size = new System.Drawing.Size(6, 23);
			// 
			// toolStripSeparator7
			// 
			this.toolStripSeparator7.AutoSize = false;
			this.toolStripSeparator7.Name = "toolStripSeparator7";
			this.toolStripSeparator7.Size = new System.Drawing.Size(6, 43);
			// 
			// btnExcel
			// 
			this.btnExcel.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(40, 37);
			this.btnExcel.Tag = "Staff_Excel";
			this.btnExcel.Text = "Excel";
			this.btnExcel.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// toolStripSeparator5
			// 
			this.toolStripSeparator5.Name = "toolStripSeparator5";
			this.toolStripSeparator5.Size = new System.Drawing.Size(6, 23);
			// 
			// btnQLNhom
			// 
			this.btnQLNhom.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnQLNhom.Image = ((System.Drawing.Image)(resources.GetObject("btnQLNhom.Image")));
			this.btnQLNhom.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnQLNhom.Name = "btnQLNhom";
			this.btnQLNhom.Size = new System.Drawing.Size(88, 37);
			this.btnQLNhom.Text = "Quản lý nhóm";
			this.btnQLNhom.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
			this.btnQLNhom.Click += new System.EventHandler(this.btnQLNhom_Click);
			// 
			// toolStripSeparator8
			// 
			this.toolStripSeparator8.Name = "toolStripSeparator8";
			this.toolStripSeparator8.Size = new System.Drawing.Size(6, 23);
			// 
			// grdData
			// 
			this.grdData.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.grdData.Cursor = System.Windows.Forms.Cursors.Default;
			this.grdData.Location = new System.Drawing.Point(3, 44);
			this.grdData.MainView = this.grvData;
			this.grdData.Name = "grdData";
			this.grdData.RepositoryItems.AddRange(new DevExpress.XtraEditors.Repository.RepositoryItem[] {
            this.NhomPhong});
			this.grdData.Size = new System.Drawing.Size(1245, 405);
			this.grdData.TabIndex = 15;
			this.grdData.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.grvData});
			this.grdData.DoubleClick += new System.EventHandler(this.grdData_DoubleClick);
			// 
			// grvData
			// 
			this.grvData.Appearance.FocusedRow.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(192)))), ((int)(((byte)(255)))));
			this.grvData.Appearance.FocusedRow.BackColor2 = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(192)))), ((int)(((byte)(255)))));
			this.grvData.Appearance.FocusedRow.Options.UseBackColor = true;
			this.grvData.Appearance.HideSelectionRow.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(192)))), ((int)(((byte)(255)))));
			this.grvData.Appearance.HideSelectionRow.BackColor2 = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(192)))), ((int)(((byte)(255)))));
			this.grvData.Appearance.HideSelectionRow.Options.UseBackColor = true;
			this.grvData.Appearance.SelectedRow.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(192)))), ((int)(((byte)(255)))));
			this.grvData.Appearance.SelectedRow.BackColor2 = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(192)))), ((int)(((byte)(255)))));
			this.grvData.Appearance.SelectedRow.Options.UseBackColor = true;
			this.grvData.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
            this.colID,
            this.colCode,
            this.colFullName,
            this.colBirthOfDate,
            this.colHomeAddress,
            this.colQualifications,
            this.colHandPhone,
            this.colMST,
            this.colBHYT,
            this.colBankAccount,
            this.colSex,
            this.colDepartmentID,
            this.colDepartmentName,
            this.colGroup,
            this.colLoginName,
            this.colStatus,
            this.colEmail});
			this.grvData.DetailHeight = 284;
			this.grvData.GridControl = this.grdData;
			this.grvData.Name = "grvData";
			this.grvData.OptionsBehavior.AutoExpandAllGroups = true;
			this.grvData.OptionsBehavior.Editable = false;
			this.grvData.OptionsFind.AlwaysVisible = true;
			this.grvData.OptionsView.ShowGroupPanel = false;
			this.grvData.CustomDrawCell += new DevExpress.XtraGrid.Views.Base.RowCellCustomDrawEventHandler(this.grvData_CustomDrawCell);
			// 
			// colID
			// 
			this.colID.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colID.AppearanceHeader.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.colID.AppearanceHeader.Options.UseBackColor = true;
			this.colID.AppearanceHeader.Options.UseFont = true;
			this.colID.AppearanceHeader.Options.UseTextOptions = true;
			this.colID.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colID.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colID.Caption = "ID";
			this.colID.FieldName = "ID";
			this.colID.MinWidth = 15;
			this.colID.Name = "colID";
			this.colID.Width = 56;
			// 
			// colCode
			// 
			this.colCode.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colCode.AppearanceHeader.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.colCode.AppearanceHeader.Options.UseBackColor = true;
			this.colCode.AppearanceHeader.Options.UseFont = true;
			this.colCode.AppearanceHeader.Options.UseTextOptions = true;
			this.colCode.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colCode.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colCode.Caption = "Mã nhân viên";
			this.colCode.FieldName = "Code";
			this.colCode.MinWidth = 15;
			this.colCode.Name = "colCode";
			this.colCode.Visible = true;
			this.colCode.VisibleIndex = 0;
			this.colCode.Width = 166;
			// 
			// colFullName
			// 
			this.colFullName.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colFullName.AppearanceHeader.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.colFullName.AppearanceHeader.Options.UseBackColor = true;
			this.colFullName.AppearanceHeader.Options.UseFont = true;
			this.colFullName.AppearanceHeader.Options.UseTextOptions = true;
			this.colFullName.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colFullName.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colFullName.Caption = "Họ tên";
			this.colFullName.FieldName = "FullName";
			this.colFullName.MinWidth = 15;
			this.colFullName.Name = "colFullName";
			this.colFullName.Visible = true;
			this.colFullName.VisibleIndex = 1;
			this.colFullName.Width = 166;
			// 
			// colBirthOfDate
			// 
			this.colBirthOfDate.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colBirthOfDate.AppearanceHeader.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.colBirthOfDate.AppearanceHeader.Options.UseBackColor = true;
			this.colBirthOfDate.AppearanceHeader.Options.UseFont = true;
			this.colBirthOfDate.AppearanceHeader.Options.UseTextOptions = true;
			this.colBirthOfDate.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colBirthOfDate.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colBirthOfDate.Caption = "Ngày sinh";
			this.colBirthOfDate.FieldName = "BirthOfDate";
			this.colBirthOfDate.MinWidth = 15;
			this.colBirthOfDate.Name = "colBirthOfDate";
			this.colBirthOfDate.Visible = true;
			this.colBirthOfDate.VisibleIndex = 4;
			this.colBirthOfDate.Width = 172;
			// 
			// colHomeAddress
			// 
			this.colHomeAddress.AppearanceCell.Options.UseTextOptions = true;
			this.colHomeAddress.AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colHomeAddress.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colHomeAddress.AppearanceHeader.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.colHomeAddress.AppearanceHeader.Options.UseBackColor = true;
			this.colHomeAddress.AppearanceHeader.Options.UseFont = true;
			this.colHomeAddress.AppearanceHeader.Options.UseTextOptions = true;
			this.colHomeAddress.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colHomeAddress.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colHomeAddress.Caption = "Địa chỉ";
			this.colHomeAddress.FieldName = "HomeAddress";
			this.colHomeAddress.MinWidth = 15;
			this.colHomeAddress.Name = "colHomeAddress";
			this.colHomeAddress.Width = 56;
			// 
			// colQualifications
			// 
			this.colQualifications.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colQualifications.AppearanceHeader.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.colQualifications.AppearanceHeader.Options.UseBackColor = true;
			this.colQualifications.AppearanceHeader.Options.UseFont = true;
			this.colQualifications.AppearanceHeader.Options.UseTextOptions = true;
			this.colQualifications.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colQualifications.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colQualifications.Caption = "Trình độ";
			this.colQualifications.FieldName = "Qualifications";
			this.colQualifications.MinWidth = 15;
			this.colQualifications.Name = "colQualifications";
			this.colQualifications.Width = 163;
			// 
			// colHandPhone
			// 
			this.colHandPhone.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colHandPhone.AppearanceHeader.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.colHandPhone.AppearanceHeader.Options.UseBackColor = true;
			this.colHandPhone.AppearanceHeader.Options.UseFont = true;
			this.colHandPhone.AppearanceHeader.Options.UseTextOptions = true;
			this.colHandPhone.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colHandPhone.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colHandPhone.Caption = "Số ĐT";
			this.colHandPhone.FieldName = "HandPhone";
			this.colHandPhone.MinWidth = 15;
			this.colHandPhone.Name = "colHandPhone";
			this.colHandPhone.Width = 76;
			// 
			// colMST
			// 
			this.colMST.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colMST.AppearanceHeader.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.colMST.AppearanceHeader.Options.UseBackColor = true;
			this.colMST.AppearanceHeader.Options.UseFont = true;
			this.colMST.AppearanceHeader.Options.UseTextOptions = true;
			this.colMST.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colMST.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colMST.Caption = "Mã số thuế TNCN";
			this.colMST.FieldName = "MST";
			this.colMST.MinWidth = 15;
			this.colMST.Name = "colMST";
			this.colMST.Width = 56;
			// 
			// colBHYT
			// 
			this.colBHYT.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colBHYT.AppearanceHeader.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.colBHYT.AppearanceHeader.Options.UseBackColor = true;
			this.colBHYT.AppearanceHeader.Options.UseFont = true;
			this.colBHYT.AppearanceHeader.Options.UseTextOptions = true;
			this.colBHYT.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colBHYT.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colBHYT.Caption = "Số thẻ BHYT";
			this.colBHYT.FieldName = "BHYT";
			this.colBHYT.MinWidth = 15;
			this.colBHYT.Name = "colBHYT";
			this.colBHYT.Width = 56;
			// 
			// colBankAccount
			// 
			this.colBankAccount.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colBankAccount.AppearanceHeader.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.colBankAccount.AppearanceHeader.Options.UseBackColor = true;
			this.colBankAccount.AppearanceHeader.Options.UseFont = true;
			this.colBankAccount.AppearanceHeader.Options.UseTextOptions = true;
			this.colBankAccount.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colBankAccount.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colBankAccount.Caption = "Số tài khoản MB";
			this.colBankAccount.FieldName = "BankAccount";
			this.colBankAccount.MinWidth = 15;
			this.colBankAccount.Name = "colBankAccount";
			this.colBankAccount.Width = 76;
			// 
			// colSex
			// 
			this.colSex.AppearanceCell.Options.UseTextOptions = true;
			this.colSex.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colSex.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.colSex.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colSex.AppearanceHeader.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.colSex.AppearanceHeader.Options.UseBackColor = true;
			this.colSex.AppearanceHeader.Options.UseFont = true;
			this.colSex.AppearanceHeader.Options.UseTextOptions = true;
			this.colSex.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colSex.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colSex.Caption = "Giới tính";
			this.colSex.FieldName = "GioiTinh";
			this.colSex.MinWidth = 15;
			this.colSex.Name = "colSex";
			this.colSex.Width = 76;
			// 
			// colDepartmentID
			// 
			this.colDepartmentID.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colDepartmentID.AppearanceHeader.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.colDepartmentID.AppearanceHeader.Options.UseBackColor = true;
			this.colDepartmentID.AppearanceHeader.Options.UseFont = true;
			this.colDepartmentID.AppearanceHeader.Options.UseTextOptions = true;
			this.colDepartmentID.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colDepartmentID.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colDepartmentID.Caption = "DepartmentID";
			this.colDepartmentID.FieldName = "DepartmentID";
			this.colDepartmentID.MinWidth = 15;
			this.colDepartmentID.Name = "colDepartmentID";
			this.colDepartmentID.Width = 56;
			// 
			// colDepartmentName
			// 
			this.colDepartmentName.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colDepartmentName.AppearanceHeader.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.colDepartmentName.AppearanceHeader.Options.UseBackColor = true;
			this.colDepartmentName.AppearanceHeader.Options.UseFont = true;
			this.colDepartmentName.AppearanceHeader.Options.UseTextOptions = true;
			this.colDepartmentName.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colDepartmentName.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colDepartmentName.Caption = "Phòng ban";
			this.colDepartmentName.FieldName = "DepartmentName";
			this.colDepartmentName.MinWidth = 15;
			this.colDepartmentName.Name = "colDepartmentName";
			this.colDepartmentName.Width = 56;
			// 
			// colGroup
			// 
			this.colGroup.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colGroup.AppearanceHeader.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.colGroup.AppearanceHeader.Options.UseBackColor = true;
			this.colGroup.AppearanceHeader.Options.UseFont = true;
			this.colGroup.AppearanceHeader.Options.UseTextOptions = true;
			this.colGroup.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colGroup.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colGroup.Caption = "Nhóm";
			this.colGroup.ColumnEdit = this.NhomPhong;
			this.colGroup.FieldName = "UserGroupID";
			this.colGroup.MinWidth = 15;
			this.colGroup.Name = "colGroup";
			this.colGroup.Width = 56;
			// 
			// NhomPhong
			// 
			this.NhomPhong.AutoHeight = false;
			this.NhomPhong.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.NhomPhong.DisplayMember = "Name";
			this.NhomPhong.Name = "NhomPhong";
			this.NhomPhong.ValueMember = "ID";
			// 
			// colLoginName
			// 
			this.colLoginName.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colLoginName.AppearanceHeader.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.colLoginName.AppearanceHeader.Options.UseBackColor = true;
			this.colLoginName.AppearanceHeader.Options.UseFont = true;
			this.colLoginName.AppearanceHeader.Options.UseTextOptions = true;
			this.colLoginName.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colLoginName.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colLoginName.Caption = "Tên đăng nhập";
			this.colLoginName.FieldName = "LoginName";
			this.colLoginName.MinWidth = 15;
			this.colLoginName.Name = "colLoginName";
			this.colLoginName.Visible = true;
			this.colLoginName.VisibleIndex = 2;
			this.colLoginName.Width = 199;
			// 
			// colStatus
			// 
			this.colStatus.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colStatus.AppearanceHeader.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.colStatus.AppearanceHeader.Options.UseBackColor = true;
			this.colStatus.AppearanceHeader.Options.UseFont = true;
			this.colStatus.AppearanceHeader.Options.UseTextOptions = true;
			this.colStatus.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colStatus.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colStatus.Caption = "Status";
			this.colStatus.FieldName = "Status";
			this.colStatus.MinWidth = 15;
			this.colStatus.Name = "colStatus";
			this.colStatus.Width = 56;
			// 
			// colEmail
			// 
			this.colEmail.AppearanceHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.colEmail.AppearanceHeader.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.colEmail.AppearanceHeader.Options.UseBackColor = true;
			this.colEmail.AppearanceHeader.Options.UseFont = true;
			this.colEmail.AppearanceHeader.Options.UseTextOptions = true;
			this.colEmail.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.colEmail.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
			this.colEmail.Caption = "Email";
			this.colEmail.FieldName = "Email";
			this.colEmail.MinWidth = 15;
			this.colEmail.Name = "colEmail";
			this.colEmail.Visible = true;
			this.colEmail.VisibleIndex = 3;
			this.colEmail.Width = 217;
			// 
			// label2
			// 
			this.label2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(1120, 63);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(102, 13);
			this.label2.TabIndex = 21;
			this.label2.Text = "Nhân viên nghỉ việc";
			// 
			// label1
			// 
			this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label1.BackColor = System.Drawing.Color.Yellow;
			this.label1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.label1.Location = new System.Drawing.Point(1079, 62);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(35, 15);
			this.label1.TabIndex = 20;
			// 
			// frmStaffManager
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(1248, 453);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.grdData);
			this.Controls.Add(this.mnuMenu);
			this.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
			this.Name = "frmStaffManager";
			this.Text = "QUẢN LÝ NHÂN SỰ";
			this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
			this.Load += new System.EventHandler(this.frmStaffManager_Load);
			this.mnuMenu.ResumeLayout(false);
			this.mnuMenu.PerformLayout();
			((System.ComponentModel.ISupportInitialize)(this.grdData)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grvData)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.NhomPhong)).EndInit();
			this.ResumeLayout(false);
			this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ToolStrip mnuMenu;
        private System.Windows.Forms.ToolStripButton btnEdit;
        private System.Windows.Forms.ToolStripButton btnDelete;
        private DevExpress.XtraGrid.GridControl grdData;
        private DevExpress.XtraGrid.Views.Grid.GridView grvData;
        private DevExpress.XtraGrid.Columns.GridColumn colID;
        private DevExpress.XtraGrid.Columns.GridColumn colCode;
        private DevExpress.XtraGrid.Columns.GridColumn colBirthOfDate;
        private DevExpress.XtraGrid.Columns.GridColumn colHomeAddress;
        private DevExpress.XtraGrid.Columns.GridColumn colQualifications;
        private DevExpress.XtraGrid.Columns.GridColumn colHandPhone;
        private DevExpress.XtraGrid.Columns.GridColumn colMST;
        private DevExpress.XtraGrid.Columns.GridColumn colBHYT;
        private DevExpress.XtraGrid.Columns.GridColumn colBankAccount;
        private DevExpress.XtraGrid.Columns.GridColumn colSex;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator7;
        private System.Windows.Forms.ToolStripButton btnExcel;
        private System.Windows.Forms.ToolStripButton btnNew;
        private DevExpress.XtraGrid.Columns.GridColumn colDepartmentID;
        private DevExpress.XtraGrid.Columns.GridColumn colDepartmentName;
        private DevExpress.XtraGrid.Columns.GridColumn colFullName;
        private DevExpress.XtraGrid.Columns.GridColumn colGroup;
        private DevExpress.XtraEditors.Repository.RepositoryItemLookUpEdit NhomPhong;
        private System.Windows.Forms.ToolStripButton btnQLNhom;
        private DevExpress.XtraGrid.Columns.GridColumn colLoginName;
        private DevExpress.XtraGrid.Columns.GridColumn colStatus;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private DevExpress.XtraGrid.Columns.GridColumn colEmail;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator4;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator5;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator8;
    }
}