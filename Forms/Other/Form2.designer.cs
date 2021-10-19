namespace BMS
{
    partial class Form2
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
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.btnTestGridBackground = new System.Windows.Forms.Button();
            this.btnErrorExcel = new System.Windows.Forms.Button();
            this.btnForm1 = new DevExpress.XtraEditors.SimpleButton();
            this.treeData = new DevExpress.XtraTreeList.TreeList();
            this.colIDTree = new DevExpress.XtraTreeList.Columns.TreeListColumn();
            this.colNameTree = new DevExpress.XtraTreeList.Columns.TreeListColumn();
            this.repositoryItemMemoEdit2 = new DevExpress.XtraEditors.Repository.RepositoryItemMemoEdit();
            this.colCodeTree = new DevExpress.XtraTreeList.Columns.TreeListColumn();
            this.button3 = new System.Windows.Forms.Button();
            this.btnTest = new System.Windows.Forms.Button();
            this.btnDownloadDMVT = new System.Windows.Forms.Button();
            this.gridControl1 = new DevExpress.XtraGrid.GridControl();
            this.gridView1 = new DevExpress.XtraGrid.Views.Grid.GridView();
            this.gridColumn1 = new DevExpress.XtraGrid.Columns.GridColumn();
            this.gridColumn2 = new DevExpress.XtraGrid.Columns.GridColumn();
            this.gridColumn3 = new DevExpress.XtraGrid.Columns.GridColumn();
            this.gridColumn4 = new DevExpress.XtraGrid.Columns.GridColumn();
            this.gridColumn5 = new DevExpress.XtraGrid.Columns.GridColumn();
            this.gridColumn6 = new DevExpress.XtraGrid.Columns.GridColumn();
            this.gridColumn7 = new DevExpress.XtraGrid.Columns.GridColumn();
            this.repositoryItemMemoEdit1 = new DevExpress.XtraEditors.Repository.RepositoryItemMemoEdit();
            this.btnBaoGia = new System.Windows.Forms.Button();
            this.btnExportExcel = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.btnGetBarcode = new System.Windows.Forms.Button();
            this.btnGetTextInImage = new System.Windows.Forms.Button();
            this.btnLoadProductModule = new System.Windows.Forms.Button();
            this.btnGetListNoPrice = new System.Windows.Forms.Button();
            this.btnAddMaterialModuleLink = new System.Windows.Forms.Button();
            this.listView1 = new System.Windows.Forms.ListView();
            this.gridControl2 = new DevExpress.XtraGrid.GridControl();
            this.gridView2 = new DevExpress.XtraGrid.Views.Grid.GridView();
            ((System.ComponentModel.ISupportInitialize)(this.treeData)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.repositoryItemMemoEdit2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridControl1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.repositoryItemMemoEdit1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridControl2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridView2)).BeginInit();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(12, 12);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(172, 23);
            this.button1.TabIndex = 0;
            this.button1.Text = "Add ảnh vào excel";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button2.Location = new System.Drawing.Point(12, 55);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(225, 23);
            this.button2.TabIndex = 1;
            this.button2.Text = "Import to database by excel";
            this.button2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // btnTestGridBackground
            // 
            this.btnTestGridBackground.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnTestGridBackground.Location = new System.Drawing.Point(12, 94);
            this.btnTestGridBackground.Name = "btnTestGridBackground";
            this.btnTestGridBackground.Size = new System.Drawing.Size(225, 23);
            this.btnTestGridBackground.TabIndex = 2;
            this.btnTestGridBackground.Text = "Test tô màu ô trong grid";
            this.btnTestGridBackground.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnTestGridBackground.UseVisualStyleBackColor = true;
            this.btnTestGridBackground.Click += new System.EventHandler(this.btnTestGridBackground_Click);
            // 
            // btnErrorExcel
            // 
            this.btnErrorExcel.Location = new System.Drawing.Point(12, 133);
            this.btnErrorExcel.Name = "btnErrorExcel";
            this.btnErrorExcel.Size = new System.Drawing.Size(119, 38);
            this.btnErrorExcel.TabIndex = 3;
            this.btnErrorExcel.Text = "Thêm lỗi từ excel";
            this.btnErrorExcel.UseVisualStyleBackColor = true;
            this.btnErrorExcel.Click += new System.EventHandler(this.btnErrorExcel_Click);
            // 
            // btnForm1
            // 
            this.btnForm1.Location = new System.Drawing.Point(243, 12);
            this.btnForm1.Name = "btnForm1";
            this.btnForm1.Size = new System.Drawing.Size(75, 76);
            this.btnForm1.TabIndex = 4;
            this.btnForm1.Text = "Form1";
            this.btnForm1.Click += new System.EventHandler(this.btnForm1_Click);
            // 
            // treeData
            // 
            this.treeData.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.treeData.Appearance.FocusedRow.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(128)))), ((int)(((byte)(255)))));
            this.treeData.Appearance.FocusedRow.BackColor2 = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(128)))), ((int)(((byte)(255)))));
            this.treeData.Appearance.FocusedRow.Options.UseBackColor = true;
            this.treeData.Appearance.HideSelectionRow.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(128)))), ((int)(((byte)(255)))));
            this.treeData.Appearance.HideSelectionRow.BackColor2 = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(128)))), ((int)(((byte)(255)))));
            this.treeData.Appearance.HideSelectionRow.Options.UseBackColor = true;
            this.treeData.Columns.AddRange(new DevExpress.XtraTreeList.Columns.TreeListColumn[] {
            this.colIDTree,
            this.colNameTree,
            this.colCodeTree});
            this.treeData.Location = new System.Drawing.Point(324, 12);
            this.treeData.Name = "treeData";
            this.treeData.OptionsBehavior.AllowExpandOnDblClick = false;
            this.treeData.OptionsBehavior.DragNodes = true;
            this.treeData.OptionsBehavior.Editable = false;
            this.treeData.ParentFieldName = "ManageId";
            this.treeData.RepositoryItems.AddRange(new DevExpress.XtraEditors.Repository.RepositoryItem[] {
            this.repositoryItemMemoEdit2});
            this.treeData.Size = new System.Drawing.Size(817, 207);
            this.treeData.TabIndex = 18;
            this.treeData.CellValueChanged += new DevExpress.XtraTreeList.CellValueChangedEventHandler(this.treeData_CellValueChanged);
            // 
            // colIDTree
            // 
            this.colIDTree.Caption = "Mã nhóm";
            this.colIDTree.FieldName = "PProductId";
            this.colIDTree.Name = "colIDTree";
            // 
            // colNameTree
            // 
            this.colNameTree.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
            this.colNameTree.AppearanceHeader.Options.UseFont = true;
            this.colNameTree.AppearanceHeader.Options.UseTextOptions = true;
            this.colNameTree.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.colNameTree.Caption = "Tên nhóm";
            this.colNameTree.ColumnEdit = this.repositoryItemMemoEdit2;
            this.colNameTree.FieldName = "TenThietBi";
            this.colNameTree.Name = "colNameTree";
            this.colNameTree.OptionsColumn.AllowEdit = false;
            this.colNameTree.OptionsColumn.AllowFocus = false;
            this.colNameTree.Visible = true;
            this.colNameTree.VisibleIndex = 1;
            this.colNameTree.Width = 396;
            // 
            // repositoryItemMemoEdit2
            // 
            this.repositoryItemMemoEdit2.Name = "repositoryItemMemoEdit2";
            // 
            // colCodeTree
            // 
            this.colCodeTree.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
            this.colCodeTree.AppearanceHeader.Options.UseFont = true;
            this.colCodeTree.AppearanceHeader.Options.UseTextOptions = true;
            this.colCodeTree.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.colCodeTree.Caption = "Mã nhóm";
            this.colCodeTree.FieldName = "MaThietBi";
            this.colCodeTree.Name = "colCodeTree";
            this.colCodeTree.OptionsColumn.AllowEdit = false;
            this.colCodeTree.OptionsColumn.AllowFocus = false;
            this.colCodeTree.Visible = true;
            this.colCodeTree.VisibleIndex = 0;
            this.colCodeTree.Width = 155;
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(12, 188);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(248, 23);
            this.button3.TabIndex = 19;
            this.button3.Text = "Download phiên bản đầu tiên của module";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // btnTest
            // 
            this.btnTest.Location = new System.Drawing.Point(12, 239);
            this.btnTest.Name = "btnTest";
            this.btnTest.Size = new System.Drawing.Size(119, 23);
            this.btnTest.TabIndex = 20;
            this.btnTest.Text = "Test";
            this.btnTest.UseVisualStyleBackColor = true;
            this.btnTest.Click += new System.EventHandler(this.btnTest_Click);
            // 
            // btnDownloadDMVT
            // 
            this.btnDownloadDMVT.Location = new System.Drawing.Point(13, 281);
            this.btnDownloadDMVT.Name = "btnDownloadDMVT";
            this.btnDownloadDMVT.Size = new System.Drawing.Size(247, 23);
            this.btnDownloadDMVT.TabIndex = 21;
            this.btnDownloadDMVT.Text = "Download danh muc vat tu cua thiet ke";
            this.btnDownloadDMVT.UseVisualStyleBackColor = true;
            this.btnDownloadDMVT.Click += new System.EventHandler(this.btnDownloadDMVT_Click);
            // 
            // gridControl1
            // 
            this.gridControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.gridControl1.Location = new System.Drawing.Point(504, 254);
            this.gridControl1.MainView = this.gridView1;
            this.gridControl1.Name = "gridControl1";
            this.gridControl1.RepositoryItems.AddRange(new DevExpress.XtraEditors.Repository.RepositoryItem[] {
            this.repositoryItemMemoEdit1});
            this.gridControl1.Size = new System.Drawing.Size(637, 64);
            this.gridControl1.TabIndex = 22;
            this.gridControl1.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.gridView1});
            // 
            // gridView1
            // 
            this.gridView1.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
            this.gridColumn1,
            this.gridColumn2,
            this.gridColumn3,
            this.gridColumn4,
            this.gridColumn5,
            this.gridColumn6,
            this.gridColumn7});
            this.gridView1.GridControl = this.gridControl1;
            this.gridView1.GroupCount = 1;
            this.gridView1.Name = "gridView1";
            this.gridView1.OptionsBehavior.AutoExpandAllGroups = true;
            this.gridView1.OptionsView.ColumnAutoWidth = false;
            this.gridView1.OptionsView.RowAutoHeight = true;
            this.gridView1.OptionsView.ShowGroupPanel = false;
            this.gridView1.SortInfo.AddRange(new DevExpress.XtraGrid.Columns.GridColumnSortInfo[] {
            new DevExpress.XtraGrid.Columns.GridColumnSortInfo(this.gridColumn1, DevExpress.Data.ColumnSortOrder.Ascending)});
            this.gridView1.DoubleClick += new System.EventHandler(this.gridView1_DoubleClick);
            // 
            // gridColumn1
            // 
            this.gridColumn1.Caption = "Module";
            this.gridColumn1.FieldName = "ModuleCode";
            this.gridColumn1.Name = "gridColumn1";
            this.gridColumn1.Visible = true;
            this.gridColumn1.VisibleIndex = 0;
            // 
            // gridColumn2
            // 
            this.gridColumn2.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
            this.gridColumn2.AppearanceHeader.Options.UseFont = true;
            this.gridColumn2.AppearanceHeader.Options.UseTextOptions = true;
            this.gridColumn2.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.gridColumn2.Caption = "STT";
            this.gridColumn2.FieldName = "STT";
            this.gridColumn2.Name = "gridColumn2";
            this.gridColumn2.Visible = true;
            this.gridColumn2.VisibleIndex = 0;
            // 
            // gridColumn3
            // 
            this.gridColumn3.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
            this.gridColumn3.AppearanceHeader.Options.UseFont = true;
            this.gridColumn3.AppearanceHeader.Options.UseTextOptions = true;
            this.gridColumn3.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.gridColumn3.Caption = "Mã vật tư";
            this.gridColumn3.FieldName = "MaterialCode";
            this.gridColumn3.Name = "gridColumn3";
            this.gridColumn3.Visible = true;
            this.gridColumn3.VisibleIndex = 1;
            // 
            // gridColumn4
            // 
            this.gridColumn4.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
            this.gridColumn4.AppearanceHeader.Options.UseFont = true;
            this.gridColumn4.AppearanceHeader.Options.UseTextOptions = true;
            this.gridColumn4.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.gridColumn4.Caption = "Tên vật tư";
            this.gridColumn4.FieldName = "MaterialName";
            this.gridColumn4.Name = "gridColumn4";
            this.gridColumn4.Visible = true;
            this.gridColumn4.VisibleIndex = 2;
            // 
            // gridColumn5
            // 
            this.gridColumn5.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
            this.gridColumn5.AppearanceHeader.Options.UseFont = true;
            this.gridColumn5.AppearanceHeader.Options.UseTextOptions = true;
            this.gridColumn5.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.gridColumn5.Caption = "Mã vật liệu";
            this.gridColumn5.FieldName = "MaVatLieu";
            this.gridColumn5.Name = "gridColumn5";
            this.gridColumn5.Visible = true;
            this.gridColumn5.VisibleIndex = 3;
            // 
            // gridColumn6
            // 
            this.gridColumn6.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
            this.gridColumn6.AppearanceHeader.Options.UseFont = true;
            this.gridColumn6.AppearanceHeader.Options.UseTextOptions = true;
            this.gridColumn6.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.gridColumn6.Caption = "Vật liệu";
            this.gridColumn6.FieldName = "VatLieu";
            this.gridColumn6.Name = "gridColumn6";
            this.gridColumn6.Visible = true;
            this.gridColumn6.VisibleIndex = 4;
            // 
            // gridColumn7
            // 
            this.gridColumn7.AppearanceHeader.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
            this.gridColumn7.AppearanceHeader.Options.UseFont = true;
            this.gridColumn7.AppearanceHeader.Options.UseTextOptions = true;
            this.gridColumn7.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.gridColumn7.Caption = "Lỗi";
            this.gridColumn7.FieldName = "Error";
            this.gridColumn7.Name = "gridColumn7";
            this.gridColumn7.Visible = true;
            this.gridColumn7.VisibleIndex = 5;
            // 
            // repositoryItemMemoEdit1
            // 
            this.repositoryItemMemoEdit1.Name = "repositoryItemMemoEdit1";
            // 
            // btnBaoGia
            // 
            this.btnBaoGia.Location = new System.Drawing.Point(12, 331);
            this.btnBaoGia.Name = "btnBaoGia";
            this.btnBaoGia.Size = new System.Drawing.Size(247, 23);
            this.btnBaoGia.TabIndex = 21;
            this.btnBaoGia.Text = "Báo giá module theo dự án";
            this.btnBaoGia.UseVisualStyleBackColor = true;
            this.btnBaoGia.Click += new System.EventHandler(this.btnBaoGia_Click);
            // 
            // btnExportExcel
            // 
            this.btnExportExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnExportExcel.Location = new System.Drawing.Point(1066, 225);
            this.btnExportExcel.Name = "btnExportExcel";
            this.btnExportExcel.Size = new System.Drawing.Size(75, 23);
            this.btnExportExcel.TabIndex = 23;
            this.btnExportExcel.Text = "Export Excel";
            this.btnExportExcel.UseVisualStyleBackColor = true;
            this.btnExportExcel.Click += new System.EventHandler(this.btnExportExcel_Click);
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(289, 310);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.textBox1.Size = new System.Drawing.Size(201, 167);
            this.textBox1.TabIndex = 25;
            // 
            // btnGetBarcode
            // 
            this.btnGetBarcode.Location = new System.Drawing.Point(303, 281);
            this.btnGetBarcode.Name = "btnGetBarcode";
            this.btnGetBarcode.Size = new System.Drawing.Size(167, 23);
            this.btnGetBarcode.TabIndex = 24;
            this.btnGetBarcode.Text = "Lấy mã vạch của bản vẽ";
            this.btnGetBarcode.UseVisualStyleBackColor = true;
            this.btnGetBarcode.Click += new System.EventHandler(this.btnGetBarcode_Click);
            // 
            // btnGetTextInImage
            // 
            this.btnGetTextInImage.Location = new System.Drawing.Point(13, 377);
            this.btnGetTextInImage.Name = "btnGetTextInImage";
            this.btnGetTextInImage.Size = new System.Drawing.Size(247, 23);
            this.btnGetTextInImage.TabIndex = 26;
            this.btnGetTextInImage.Text = "Lấy ký tự trong hình ảnh";
            this.btnGetTextInImage.UseVisualStyleBackColor = true;
            this.btnGetTextInImage.Click += new System.EventHandler(this.btnGetTextInImage_Click);
            // 
            // btnLoadProductModule
            // 
            this.btnLoadProductModule.Location = new System.Drawing.Point(324, 221);
            this.btnLoadProductModule.Name = "btnLoadProductModule";
            this.btnLoadProductModule.Size = new System.Drawing.Size(144, 30);
            this.btnLoadProductModule.TabIndex = 27;
            this.btnLoadProductModule.Text = "Load module của dự án";
            this.btnLoadProductModule.UseVisualStyleBackColor = true;
            this.btnLoadProductModule.Click += new System.EventHandler(this.btnLoadProductModule_Click);
            // 
            // btnGetListNoPrice
            // 
            this.btnGetListNoPrice.Location = new System.Drawing.Point(13, 420);
            this.btnGetListNoPrice.Name = "btnGetListNoPrice";
            this.btnGetListNoPrice.Size = new System.Drawing.Size(247, 23);
            this.btnGetListNoPrice.TabIndex = 28;
            this.btnGetListNoPrice.Text = "Lấy danh sách vật tư không có giá";
            this.btnGetListNoPrice.UseVisualStyleBackColor = true;
            this.btnGetListNoPrice.Click += new System.EventHandler(this.btnGetListNoPrice_Click);
            // 
            // btnAddMaterialModuleLink
            // 
            this.btnAddMaterialModuleLink.Location = new System.Drawing.Point(12, 454);
            this.btnAddMaterialModuleLink.Name = "btnAddMaterialModuleLink";
            this.btnAddMaterialModuleLink.Size = new System.Drawing.Size(247, 23);
            this.btnAddMaterialModuleLink.TabIndex = 29;
            this.btnAddMaterialModuleLink.Text = "Thêm danh sách vật tư cho module";
            this.btnAddMaterialModuleLink.UseVisualStyleBackColor = true;
            this.btnAddMaterialModuleLink.Click += new System.EventHandler(this.btnAddMaterialModuleLink_Click);
            // 
            // listView1
            // 
            this.listView1.Location = new System.Drawing.Point(496, 324);
            this.listView1.Name = "listView1";
            this.listView1.Size = new System.Drawing.Size(149, 162);
            this.listView1.TabIndex = 30;
            this.listView1.UseCompatibleStateImageBehavior = false;
            // 
            // gridControl2
            // 
            this.gridControl2.Location = new System.Drawing.Point(651, 324);
            this.gridControl2.MainView = this.gridView2;
            this.gridControl2.Name = "gridControl2";
            this.gridControl2.Size = new System.Drawing.Size(490, 162);
            this.gridControl2.TabIndex = 31;
            this.gridControl2.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.gridView2});
            // 
            // gridView2
            // 
            this.gridView2.GridControl = this.gridControl2;
            this.gridView2.Name = "gridView2";
            // 
            // Form2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1153, 489);
            this.Controls.Add(this.gridControl2);
            this.Controls.Add(this.listView1);
            this.Controls.Add(this.btnAddMaterialModuleLink);
            this.Controls.Add(this.btnGetListNoPrice);
            this.Controls.Add(this.btnLoadProductModule);
            this.Controls.Add(this.btnGetTextInImage);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.btnGetBarcode);
            this.Controls.Add(this.btnExportExcel);
            this.Controls.Add(this.gridControl1);
            this.Controls.Add(this.btnBaoGia);
            this.Controls.Add(this.btnDownloadDMVT);
            this.Controls.Add(this.btnTest);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.treeData);
            this.Controls.Add(this.btnForm1);
            this.Controls.Add(this.btnErrorExcel);
            this.Controls.Add(this.btnTestGridBackground);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Name = "Form2";
            this.Text = "Form1";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.treeData)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.repositoryItemMemoEdit2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridControl1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.repositoryItemMemoEdit1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridControl2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridView2)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button btnTestGridBackground;
        private System.Windows.Forms.Button btnErrorExcel;
        private DevExpress.XtraEditors.SimpleButton btnForm1;
        private DevExpress.XtraTreeList.TreeList treeData;
        private DevExpress.XtraTreeList.Columns.TreeListColumn colIDTree;
        private DevExpress.XtraTreeList.Columns.TreeListColumn colNameTree;
        private DevExpress.XtraEditors.Repository.RepositoryItemMemoEdit repositoryItemMemoEdit2;
        private DevExpress.XtraTreeList.Columns.TreeListColumn colCodeTree;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button btnTest;
        private System.Windows.Forms.Button btnDownloadDMVT;
        private DevExpress.XtraGrid.GridControl gridControl1;
        private DevExpress.XtraGrid.Views.Grid.GridView gridView1;
        private DevExpress.XtraEditors.Repository.RepositoryItemMemoEdit repositoryItemMemoEdit1;
        private System.Windows.Forms.Button btnBaoGia;
        private System.Windows.Forms.Button btnExportExcel;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Button btnGetBarcode;
        private System.Windows.Forms.Button btnGetTextInImage;
        private System.Windows.Forms.Button btnLoadProductModule;
        private System.Windows.Forms.Button btnGetListNoPrice;
        private DevExpress.XtraGrid.Columns.GridColumn gridColumn1;
        private DevExpress.XtraGrid.Columns.GridColumn gridColumn2;
        private DevExpress.XtraGrid.Columns.GridColumn gridColumn3;
        private DevExpress.XtraGrid.Columns.GridColumn gridColumn4;
        private DevExpress.XtraGrid.Columns.GridColumn gridColumn5;
        private DevExpress.XtraGrid.Columns.GridColumn gridColumn6;
        private DevExpress.XtraGrid.Columns.GridColumn gridColumn7;
        private System.Windows.Forms.Button btnAddMaterialModuleLink;
        private System.Windows.Forms.ListView listView1;
        private DevExpress.XtraGrid.GridControl gridControl2;
        private DevExpress.XtraGrid.Views.Grid.GridView gridView2;
    }
}

