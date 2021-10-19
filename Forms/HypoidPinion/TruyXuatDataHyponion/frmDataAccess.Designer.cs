
namespace BMS
{
    partial class frmDataAccess
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
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmDataAccess));
			this.label1 = new System.Windows.Forms.Label();
			this.txtCode = new System.Windows.Forms.TextBox();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.txtGoods = new System.Windows.Forms.TextBox();
			this.txtOrder = new System.Windows.Forms.TextBox();
			this.xtraTabbedMdiManager1 = new DevExpress.XtraTabbedMdi.XtraTabbedMdiManager(this.components);
			this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
			this.btnMoldResult = new System.Windows.Forms.Button();
			this.btnJigCode = new System.Windows.Forms.Button();
			this.btnKnifeCode = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.xtraTabbedMdiManager1)).BeginInit();
			this.tableLayoutPanel1.SuspendLayout();
			this.SuspendLayout();
			// 
			// label1
			// 
			this.label1.BackColor = System.Drawing.Color.DeepSkyBlue;
			this.label1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.label1.Location = new System.Drawing.Point(1, 1);
			this.label1.Margin = new System.Windows.Forms.Padding(1);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(234, 44);
			this.label1.TabIndex = 0;
			this.label1.Text = "THÔNG TIN SẢN PHẨM";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// txtCode
			// 
			this.txtCode.BackColor = System.Drawing.Color.Violet;
			this.tableLayoutPanel1.SetColumnSpan(this.txtCode, 4);
			this.txtCode.Dock = System.Windows.Forms.DockStyle.Fill;
			this.txtCode.Font = new System.Drawing.Font("Microsoft Sans Serif", 24F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtCode.Location = new System.Drawing.Point(237, 1);
			this.txtCode.Margin = new System.Windows.Forms.Padding(1);
			this.txtCode.Name = "txtCode";
			this.txtCode.Size = new System.Drawing.Size(946, 44);
			this.txtCode.TabIndex = 1;
			this.txtCode.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
			this.txtCode.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCode_KeyDown);
			// 
			// label2
			// 
			this.label2.BackColor = System.Drawing.Color.RoyalBlue;
			this.label2.Dock = System.Windows.Forms.DockStyle.Fill;
			this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.label2.ForeColor = System.Drawing.Color.White;
			this.label2.Location = new System.Drawing.Point(1, 47);
			this.label2.Margin = new System.Windows.Forms.Padding(1);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(234, 44);
			this.label2.TabIndex = 2;
			this.label2.Text = "MÃ HÀNG";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label3
			// 
			this.label3.BackColor = System.Drawing.Color.RoyalBlue;
			this.label3.Dock = System.Windows.Forms.DockStyle.Fill;
			this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.label3.ForeColor = System.Drawing.Color.White;
			this.label3.Location = new System.Drawing.Point(237, 47);
			this.label3.Margin = new System.Windows.Forms.Padding(1);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(234, 44);
			this.label3.TabIndex = 3;
			this.label3.Text = "ORDER GIA CÔNG";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label4
			// 
			this.label4.BackColor = System.Drawing.Color.RoyalBlue;
			this.label4.Dock = System.Windows.Forms.DockStyle.Fill;
			this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.label4.ForeColor = System.Drawing.Color.White;
			this.label4.Location = new System.Drawing.Point(473, 47);
			this.label4.Margin = new System.Windows.Forms.Padding(1);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(234, 44);
			this.label4.TabIndex = 4;
			this.label4.Text = "MÃ DAO";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label5
			// 
			this.label5.BackColor = System.Drawing.Color.RoyalBlue;
			this.label5.Dock = System.Windows.Forms.DockStyle.Fill;
			this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.label5.ForeColor = System.Drawing.Color.White;
			this.label5.Location = new System.Drawing.Point(709, 47);
			this.label5.Margin = new System.Windows.Forms.Padding(1);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(234, 44);
			this.label5.TabIndex = 5;
			this.label5.Text = "MÃ JIG";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label6
			// 
			this.label6.BackColor = System.Drawing.Color.RoyalBlue;
			this.label6.Dock = System.Windows.Forms.DockStyle.Fill;
			this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.label6.ForeColor = System.Drawing.Color.White;
			this.label6.Location = new System.Drawing.Point(945, 47);
			this.label6.Margin = new System.Windows.Forms.Padding(1);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(238, 44);
			this.label6.TabIndex = 6;
			this.label6.Text = "KẾT QUẢ KIỂM TRA THAY KHUÔN";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// txtGoods
			// 
			this.txtGoods.Dock = System.Windows.Forms.DockStyle.Fill;
			this.txtGoods.Enabled = false;
			this.txtGoods.Font = new System.Drawing.Font("Microsoft Sans Serif", 24F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtGoods.Location = new System.Drawing.Point(1, 93);
			this.txtGoods.Margin = new System.Windows.Forms.Padding(1);
			this.txtGoods.Name = "txtGoods";
			this.txtGoods.Size = new System.Drawing.Size(234, 44);
			this.txtGoods.TabIndex = 7;
			// 
			// txtOrder
			// 
			this.txtOrder.Dock = System.Windows.Forms.DockStyle.Fill;
			this.txtOrder.Enabled = false;
			this.txtOrder.Font = new System.Drawing.Font("Microsoft Sans Serif", 24F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtOrder.Location = new System.Drawing.Point(237, 93);
			this.txtOrder.Margin = new System.Windows.Forms.Padding(1);
			this.txtOrder.Name = "txtOrder";
			this.txtOrder.Size = new System.Drawing.Size(234, 44);
			this.txtOrder.TabIndex = 8;
			// 
			// xtraTabbedMdiManager1
			// 
			this.xtraTabbedMdiManager1.ClosePageButtonShowMode = DevExpress.XtraTab.ClosePageButtonShowMode.InAllTabPageHeaders;
			this.xtraTabbedMdiManager1.MdiParent = this;
			// 
			// tableLayoutPanel1
			// 
			this.tableLayoutPanel1.ColumnCount = 5;
			this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 20F));
			this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 20F));
			this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 20F));
			this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 20F));
			this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 20F));
			this.tableLayoutPanel1.Controls.Add(this.btnMoldResult, 4, 2);
			this.tableLayoutPanel1.Controls.Add(this.btnJigCode, 3, 2);
			this.tableLayoutPanel1.Controls.Add(this.label1, 0, 0);
			this.tableLayoutPanel1.Controls.Add(this.label6, 4, 1);
			this.tableLayoutPanel1.Controls.Add(this.txtCode, 1, 0);
			this.tableLayoutPanel1.Controls.Add(this.label5, 3, 1);
			this.tableLayoutPanel1.Controls.Add(this.label2, 0, 1);
			this.tableLayoutPanel1.Controls.Add(this.txtGoods, 0, 2);
			this.tableLayoutPanel1.Controls.Add(this.label4, 2, 1);
			this.tableLayoutPanel1.Controls.Add(this.txtOrder, 1, 2);
			this.tableLayoutPanel1.Controls.Add(this.label3, 1, 1);
			this.tableLayoutPanel1.Controls.Add(this.btnKnifeCode, 2, 2);
			this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Top;
			this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
			this.tableLayoutPanel1.Name = "tableLayoutPanel1";
			this.tableLayoutPanel1.RowCount = 3;
			this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 33.33333F));
			this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 33.33333F));
			this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 33.33333F));
			this.tableLayoutPanel1.Size = new System.Drawing.Size(1184, 138);
			this.tableLayoutPanel1.TabIndex = 14;
			// 
			// btnMoldResult
			// 
			this.btnMoldResult.Dock = System.Windows.Forms.DockStyle.Fill;
			this.btnMoldResult.Font = new System.Drawing.Font("Microsoft Sans Serif", 24F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnMoldResult.Location = new System.Drawing.Point(945, 93);
			this.btnMoldResult.Margin = new System.Windows.Forms.Padding(1);
			this.btnMoldResult.Name = "btnMoldResult";
			this.btnMoldResult.Size = new System.Drawing.Size(238, 44);
			this.btnMoldResult.TabIndex = 11;
			this.btnMoldResult.UseVisualStyleBackColor = true;
			this.btnMoldResult.Click += new System.EventHandler(this.btnMoldResult_Click);
			// 
			// btnJigCode
			// 
			this.btnJigCode.Dock = System.Windows.Forms.DockStyle.Fill;
			this.btnJigCode.Font = new System.Drawing.Font("Microsoft Sans Serif", 24F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnJigCode.Location = new System.Drawing.Point(709, 93);
			this.btnJigCode.Margin = new System.Windows.Forms.Padding(1);
			this.btnJigCode.Name = "btnJigCode";
			this.btnJigCode.Size = new System.Drawing.Size(234, 44);
			this.btnJigCode.TabIndex = 10;
			this.btnJigCode.UseVisualStyleBackColor = true;
			this.btnJigCode.Click += new System.EventHandler(this.btnJigCode_Click);
			// 
			// btnKnifeCode
			// 
			this.btnKnifeCode.Dock = System.Windows.Forms.DockStyle.Fill;
			this.btnKnifeCode.Font = new System.Drawing.Font("Microsoft Sans Serif", 24F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnKnifeCode.Location = new System.Drawing.Point(473, 93);
			this.btnKnifeCode.Margin = new System.Windows.Forms.Padding(1);
			this.btnKnifeCode.Name = "btnKnifeCode";
			this.btnKnifeCode.Size = new System.Drawing.Size(234, 44);
			this.btnKnifeCode.TabIndex = 9;
			this.btnKnifeCode.UseVisualStyleBackColor = true;
			this.btnKnifeCode.Click += new System.EventHandler(this.btnKnifeCode_Click);
			// 
			// frmDataAccess
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.BackColor = System.Drawing.Color.White;
			this.ClientSize = new System.Drawing.Size(1184, 656);
			this.Controls.Add(this.tableLayoutPanel1);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.IsMdiContainer = true;
			this.MinimumSize = new System.Drawing.Size(921, 695);
			this.Name = "frmDataAccess";
			this.Text = "TRUY XUẤT DỮ LIỆU";
			this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
			this.Load += new System.EventHandler(this.frmDataAccess_Load);
			((System.ComponentModel.ISupportInitialize)(this.xtraTabbedMdiManager1)).EndInit();
			this.tableLayoutPanel1.ResumeLayout(false);
			this.tableLayoutPanel1.PerformLayout();
			this.ResumeLayout(false);

        }

		#endregion

		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.TextBox txtCode;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.TextBox txtGoods;
		private System.Windows.Forms.TextBox txtOrder;
		private DevExpress.XtraTabbedMdi.XtraTabbedMdiManager xtraTabbedMdiManager1;
		private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
		private System.Windows.Forms.Button btnMoldResult;
		private System.Windows.Forms.Button btnJigCode;
		private System.Windows.Forms.Button btnKnifeCode;
	}
}

