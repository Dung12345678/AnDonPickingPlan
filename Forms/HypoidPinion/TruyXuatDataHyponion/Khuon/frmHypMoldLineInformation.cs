using HP.Business;
using HP.Model;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BMS
{
	public partial class frmHypMoldLineInformation : _Forms
	{
		#region Variables
		public string Order = "";
		public string Goods = "";
		public string Code = "";
		int _KnifeDetailListID;
		DataTable _dt;
		DataTable _dtDisplay;
		DataTable _dtNoDisplay;
		int _stepIndex = 0;
		int _NoIndefinitely = 9999;
		string _path;
		Color _colorEmpty;

		CheckHistoryDetailModel _checkHistoryDetailModel;
		#endregion Variables
		public frmHypMoldLineInformation()
		{
			InitializeComponent();
		}
		/// <summary>
		/// Load Form
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void Form1_Load(object sender, EventArgs e)
		{
			LoadcboPart();
			loadPath();
			_colorEmpty = Color.FromArgb(255, 192, 255);
			txtWorkerCode.Focus();
			EnableText(true);
			LoadMasterAnđetails();
		}
		void LoadcboPart()
		{
			DataTable dt = TextUtilsHP.Select("SELECT * FROM [HypoidPinion].[dbo].[Part]");
			cboPart.DataSource = dt;
			cboPart.DisplayMember = "PartName";
			cboPart.ValueMember = "ID";
		}
		void LoadMasterAnđetails()
		{
			DataSet ds = TextUtilsHP.LoadDataSetFromSP("spLoadMoldMasterAndDetails", new string[] { "@Goods", "@Code", "@Order" }, new object[] { Goods, Code, Order });
			DataTable dtMaster = ds.Tables[0];
			DataTable dtDetails = ds.Tables[1];
			if (dtMaster.Rows.Count > 0)
			{
				txtWorkerCode.Text = TextUtilsHP.ToString(dtMaster.Rows[0]["WorkerCode"]);
				txtOrder.Text = TextUtilsHP.ToString(dtMaster.Rows[0]["OrderCode"]);
				cboPart.SelectedItem = TextUtilsHP.ToString(dtMaster.Rows[0]["PartID"]);
				txtGoodsCode.Text = TextUtilsHP.ToString(dtMaster.Rows[0]["GoodsCode"]);
				txtMachine.Text = TextUtilsHP.ToString(dtMaster.Rows[0]["Machine"]);
				txtKnifeDetailCode.Text = TextUtilsHP.ToString(dtMaster.Rows[0]["KnifeDetailCode"]);
				cboCongDoan.Text = TextUtilsHP.ToString(dtMaster.Rows[0]["StepCode"]);
			}
			if (dtDetails.Rows.Count > 0)
			{
				DataTable _dtNoDisplay = new DataTable();
				DataTable dtDisplay = new DataTable();
				dtDisplay = dtDetails.Clone();
				dtDisplay.Clear();
				DataRow[] dr = dtDetails.Select("isDisplay=0");
				_dtNoDisplay = dtDetails.Clone();
				_dtNoDisplay.Clear();
				foreach (DataRow item in dr)
				{
					_dtNoDisplay.ImportRow(item);
				}
				grdData.DataSource = _dtNoDisplay;
				DataRow[] drText = dtDetails.Select("isDisplay=1");
				foreach (DataRow item in drText)
				{
					dtDisplay.ImportRow(item);
				}
				for (int i = 0; i < dtDisplay.Rows.Count; i++)
				{
					int LocationX = TextUtilsHP.ToInt(dtDisplay.Rows[i]["X"]);
					int LocationY = TextUtilsHP.ToInt(dtDisplay.Rows[i]["Y"]);
					string RealValue = TextUtilsHP.ToString(dtDisplay.Rows[i]["RealValue"]);
					string Result = TextUtilsHP.ToString(dtDisplay.Rows[i]["Result"]);
					showTextBox(LocationX, LocationY, RealValue, i, Result);
				}
			}
			try
			{
				string path = _path + $"\\{txtKnifeDetailCode.Text.Trim()}" + ".pdf";
				pdfViewer1.LoadDocument(path);
			}
			catch
			{
				MessageBox.Show("Không tìm thấy file PDF", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				return;
			}
		}
		void EnableText(bool check)
		{
			cboPart.Enabled = check;
			txtMachine.Enabled = check;
			txtKnifeDetailCode.Enabled = check;
			txtOrder.Enabled = check;
			txtWorkerCode.Enabled = check;
			cboCongDoan.Enabled = check;
		}
		private void loadPath()
		{
			if (File.Exists(Application.StartupPath + @"\path.txt"))
			{
				_path = File.ReadAllText(Application.StartupPath + @"\path.txt");
			}
		}
		private void showTextBox(int X, int Y, string RealValue, int row, string Result)
		{
			//TODO: Tạo textbox
			TextBox txtRow = new TextBox();
			txtRow.Name = "txtxxx" + $"{row}";
			txtRow.Font = new Font("Calibri", 10);
			txtRow.Size = new Size(50, 10);
			txtRow.Location = new Point(X, Y);
			txtRow.Dock = DockStyle.None;
			txtRow.BackColor = Color.FromArgb(255, 192, 255);
			txtRow.ForeColor = Color.Black;
			txtRow.Enabled = true;
			txtRow.BringToFront();
			txtRow.Enabled = true;
			txtRow.Text = RealValue;
			panel2.Controls.Add(txtRow);
			pdfViewer1.SendToBack();
			if (Result != "OK")
			{
				txtRow.BackColor = Color.Red;
			}
			else
			{
				txtRow.BackColor = Color.Lime;
			}
			//txtRow.TextChanged += new System.EventHandler(this.Textbox_TextChanged);
			//txtRow.KeyDown += new KeyEventHandler(Textbox_KeyDown);
		//	txtRow.Click += new EventHandler(Textbox_Click);
		}

		private void btnPath_Click(object sender, EventArgs e)
		{
			FolderBrowserDialog dlg = new FolderBrowserDialog();
			if (dlg.ShowDialog() == DialogResult.OK)
			{
				string path = dlg.SelectedPath;
				string fileSavePath = Application.StartupPath + @"\path.txt";
				if (File.Exists(fileSavePath) == false)
				{
					File.Create(fileSavePath);
				}
				this.Invoke((MethodInvoker)delegate
				{
					File.WriteAllText(fileSavePath, path);
				});

				loadPath();
			}
		}
		private void grvData_RowCellStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowCellStyleEventArgs e)
		{
			if (e.RowHandle < 0) return;
			if (e.Column.VisibleIndex < 1) return;

			if (e.Column == colRealvalue)
			{
				string Evaluate = TextUtilsHP.ToString(grvData.GetRowCellValue(e.RowHandle, colResult));
				if (Evaluate != "OK")
				{
					e.Appearance.BackColor = Color.Red;
				}
				else
				{
					e.Appearance.BackColor = Color.Lime;
				}
			}
		}
	}
}
