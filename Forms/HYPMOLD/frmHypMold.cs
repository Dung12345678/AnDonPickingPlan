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
	public partial class frmHypMold : _Forms
	{
		#region Variables
		bool _isNew = false;
		Control _activeControl;
		Point _previousPosition;
		TextBox txt;
		int _KnifeDetailListID;
		int _id;
		int IDNew = 0;
		int _oldSelectedRow;
		string _path;
		int isZoom = 0;
		#endregion Variables
		public frmHypMold()
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

			SetInterface(false);
			Setbutton(false);
			loadPath();
		}
		private void loadPath()
		{
			if (File.Exists(Application.StartupPath + @"\path.txt"))
			{
				_path = File.ReadAllText(Application.StartupPath + @"\path.txt");
			}
		}
		private void Setbutton(bool isEdit)
		{
			btnEditTxt.Enabled = btnNew.Enabled = btnEdit.Enabled = btnDelete.Enabled = btnCancel.Enabled = btnSave.Enabled = isEdit;
		}
		private void SetInterface(bool isEdit)
		{
			btnSave.Visible = isEdit;
			btnCancel.Visible = isEdit;

			btnPath.Visible = !isEdit;
			btnNew.Visible = !isEdit;
			btnEdit.Visible = !isEdit;
			btnEditTxt.Visible = !isEdit;
			btnDelete.Visible = !isEdit;
		}
		/// <summary>
		/// Sự kiện Keydown của textbox mã hàng
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void txtMaHang_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode != Keys.Enter) return;
			//Load cbo theo mã jig hoặc mã dao
			DataTable dt = TextUtilsHP.Select($"SELECT DISTINCT ProductStepCode FROM [HypoidPinion].[dbo].[KnifeStep] kw JOIN dbo.KnifeDetailList kl ON kw.KnifeDetailListID = kl.ID WHERE kl.KnifeCode = N'{txtMaHang.Text.Trim()}'");
			DataRow dr = dt.NewRow();
			dr["ProductStepCode"] = "";
			dt.Rows.InsertAt(dr, 0);
			cboCongDoan.DataSource = dt;
			cboCongDoan.DisplayMember = "ProductStepCode";

			//cboStep.ValueMember = "ID";
			cboCongDoan.Focus();

			//loadAllTextBox();
		}
		private void loadAllTextBox()
		{
			try
			{
				//TODO: Xóa hết các ô textbox cũ
				deleteAllTextBox();
				//TODO: Load Bảng mục kiểm tra theo ID của tên mã hàng
				DataTable dt = TextUtilsHP.Select($"select ID from dbo.KnifeDetailList where KnifeCode = '{txtMaHang.Text}'");
				if (dt.Rows.Count <= 0)
				{
					MessageBox.Show("Không tìm thấy mã khuôn", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
					Setbutton(false);
					return;
				}
				//TODO: Load file PDF theo tên
				try
				{
					string path = _path + $"\\{txtMaHang.Text.Trim()}.pdf";
					pdfViewer1.LoadDocument(path);
				}
				catch
				{
					MessageBox.Show("Không tìm thấy file PDF", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				}
				Setbutton(true);
				_KnifeDetailListID = TextUtilsHP.ToInt(dt.Rows[0]["ID"]);
				// Nhét data của 1 cột vào 1 list
				List<int> isDisplay = new List<int>(loadDataHypMold().Rows.Count);
				foreach (DataRow row in loadDataHypMold().Rows)
					if (row["isDisplay"] != null)
						isDisplay.Add((int)row["isDisplay"]);
				//TODO: Load các textbox 
				for (int i = 0; i < loadDataHypMold().Rows.Count; i++)
				{
					if (isDisplay[i] == 1)
					{
						int LocationX = TextUtilsHP.ToInt(loadDataHypMold().Rows[i]["X"]);
						int LocationY = TextUtilsHP.ToInt(loadDataHypMold().Rows[i]["Y"]);
						//string RealValue = TextUtilsHP.ToString(loadDataHypMold().Rows[i]["RealValue"]);
						showTextbox(LocationX, LocationY, "");
					}
				}
				setColorTable();
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.ToString());
				Setbutton(false);
			}
		}
		private void deleteAllTextBox()
		{

			foreach (Control item in panel2.Controls.OfType<TextBox>().ToList())
			{
				panel2.Controls.Remove(item);
			}

		}
		private DataTable loadDataHypMold()
		{
			DataTable dt = TextUtilsHP.GetDataTableFromSP("spLoadJig", new string[] { "@KnifeCode", "@StepCode" }, new object[] { txtMaHang.Text.Trim(), cboCongDoan.Text.Trim() });
			grvKnifeJigWorking.AutoGenerateColumns = false; // Không show tất cả cột lên
			grvKnifeJigWorking.DataSource = dt;
			return dt;
			//string sql = $"Select * from dbo.KnifeJigWorking where KnifeDetailListID ={_KnifeDetailListID} order by SortOrder ";
			//DataTable dt = TextUtilsHP.Select(sql);
			//grvKnifeJigWorking.AutoGenerateColumns = false; // Không show tất cả cột lên
			//grvKnifeJigWorking.DataSource = dt;
			//return dt;
		}
		private void showTextbox(int X, int Y, string RealValue)
		{
			TextBox txtRow = new TextBox();
			txtRow.Font = new Font("Calibri", 10);
			txtRow.Size = new Size(50, 10);
			txtRow.Location = new Point(X, Y);
			txtRow.Dock = DockStyle.None;
			txtRow.Enabled = false;
			txtRow.BringToFront();
			txtRow.Text = RealValue;
			panel2.Controls.Add(txtRow);
			pdfViewer1.SendToBack();
		}
		private void setColorTable()
		{
			for (int j = 0; j < grvKnifeJigWorking.Rows.Count; j++)
			{
				if (grvKnifeJigWorking.Rows[j].Cells["colDisplay"].Value.ToString().Contains("1"))
				{
					grvKnifeJigWorking.Rows[j].Cells["colSTT"].Style.BackColor = Color.Cyan;
				}
				//else
				//{
				//    grvKnifeJigWorking.Rows[j].Cells["colEvaluate"].Style.BackColor = Color.Violet;
				//}
			}
		}
		/// <summary>
		/// Thêm sửa xóa 
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>

		private void btnNew_Click(object sender, EventArgs e)
		{
			deleteAllTextBox();
			creatNewData();
		}
		private void creatNewData()
		{
			//TODO: Tạo data
			frmHypMoldDetail frm = new frmHypMoldDetail();
			frm._MaHang = txtMaHang.Text;
			frm._KnifeDetailListID = _KnifeDetailListID;
			frm._isNew = true;
			frm.CongDoan = cboCongDoan.Text;
			frm.Createby = txtCreateby.Text.Trim();
			if (frm.ShowDialog() == DialogResult.OK)
			{
				if (frm.KnifeID != -1)
				{
					//loadDataHypMold();
					IDNew = frm.KnifeID;
					_oldSelectedRow = grvKnifeJigWorking.RowCount - 1;
					DataRow[] dtr = loadDataHypMold().Select($"ID = {frm.KnifeID}");
					if (dtr.Length <= 0) return;
					int X = TextUtilsHP.ToInt(dtr[0]["X"]);
					int Y = TextUtilsHP.ToInt(dtr[0]["Y"]);
					if (TextUtilsHP.ToInt(dtr[0]["isDisplay"]) == 1) creatNewTextBox(X, Y, "");
					if (frm._ischeck == 1)
					{
						SetInterface(true);
						grvKnifeJigWorking.Enabled = false;
					}
					else
					{
						SetInterface(false);
						grvKnifeJigWorking.Enabled = true;
						loadAllTextBox();
						_oldSelectedRow = grvKnifeJigWorking.Rows.Count;
					}
				}
			}
		}
		private void creatNewTextBox(int X, int Y, string vaLue)
		{
			txt = new TextBox();
			txt.Font = new Font("Calibri", 10);
			txt.Size = new Size(50, 10);
			txt.Location = new Point(X, Y);
			txt.Text = vaLue;
			txt.Dock = DockStyle.None;
			txt.BringToFront();
			txt.ReadOnly = true; //Không thể chỉnh sửa text box
			txt.MouseMove += new MouseEventHandler(Control_MouseMove);
			txt.MouseUp += new MouseEventHandler(Control_MouseUp);
			txt.MouseDown += new MouseEventHandler(Control_MouseDown);
			panel2.Controls.Add(txt);
			pdfViewer1.SendToBack();
			lblx.Text = txt.Location.X.ToString();
			lbly.Text = txt.Location.Y.ToString();
		}
		private void Control_MouseMove(object sender, MouseEventArgs e)
		{
			if (_activeControl == null || _activeControl != sender)
			{
				return;
			}
			if (e.Button == MouseButtons.Left)
			{
				var location = _activeControl.Location;
				location.Offset(e.Location.X - _previousPosition.X, e.Location.Y - _previousPosition.Y);
				_activeControl.Location = location;
				lblx.Text = txt.Location.X.ToString();
				lbly.Text = txt.Location.Y.ToString();
				//_activeControl.BringToFront();
			}

		}
		private void Control_MouseUp(object sender, MouseEventArgs e)
		{
			if (_activeControl == null || _activeControl != sender)
			{
				return;
			}
			if (e.Button == MouseButtons.Left)
			{
				var location = _activeControl.Location;
				location.Offset(e.Location.X - _previousPosition.X, e.Location.Y - _previousPosition.Y);
				_activeControl.Location = location;
			}

		}
		private void Control_MouseDown(object sender, MouseEventArgs e)
		{
			_activeControl = sender as Control;
			_previousPosition = e.Location;
			Cursor = Cursors.Hand;
		}
		private void btnEditTxt_Click(object sender, EventArgs e)
		{
			try
			{
				int selectedrowindex = grvKnifeJigWorking.SelectedCells[0].RowIndex;
				_oldSelectedRow = selectedrowindex;
				int ID = TextUtilsHP.ToInt(grvKnifeJigWorking.Rows[selectedrowindex].Cells["colID"].Value);
				deleteAllTextBox();
				DataRow[] dtr = loadDataHypMold().Select($" ID = {ID}");
				if (TextUtilsHP.ToInt(dtr[0]["isDisplay"]) == 1)
				{
					grvKnifeJigWorking.Visible = false;
					pdfViewer1.Visible = true;
					int X = TextUtilsHP.ToInt(dtr[0]["X"]);
					int Y = TextUtilsHP.ToInt(dtr[0]["Y"]);
					//string RealValue = TextUtilsHP.ToString(dtr[0]["RealValue"]);
					//creatNewTextBox(X, Y, RealValue);
					creatNewTextBox(X, Y, "");
					grvKnifeJigWorking.Enabled = false;
					SetInterface(true);
				}
				else
				{
					grvKnifeJigWorking.Enabled = true;
					SetInterface(false);
					loadAllTextBox();
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.ToString());
			}

		}
		private void btnSave_Click(object sender, EventArgs e)
		{
			grvKnifeJigWorking.Visible = true;
			int ID = 0;
			if (IDNew == 0)
			{
				ID = TextUtilsHP.ToInt(grvKnifeJigWorking.Rows[_oldSelectedRow].Cells["colID"].Value);
			}
			else
			{
				ID = IDNew;
			}

			string sqlUpdate = $"update dbo.KnifeJigWorking set X = '{lblx.Text}', Y = '{lbly.Text}' where ID ={ID}";
			TextUtilsHP.ExcuteSQL(sqlUpdate);
			txt.Enabled = false;
			loadAllTextBox();
			SetInterface(false);
			grvKnifeJigWorking.Enabled = true;
			IDNew = 0;
		}
		private void btnCancel_Click(object sender, EventArgs e)
		{
			SetInterface(false);
			loadAllTextBox();
			grvKnifeJigWorking.Enabled = true;
		}
		private void btnEdit_Click(object sender, EventArgs e)
		{
			if (grvKnifeJigWorking.Rows.Count <= 0) return;
			int selectedrowindex = grvKnifeJigWorking.SelectedCells[0].RowIndex;
			_oldSelectedRow = selectedrowindex;
			int ID = TextUtilsHP.ToInt(grvKnifeJigWorking.Rows[selectedrowindex].Cells["colID"].Value);
			frmHypMoldDetail frm = new frmHypMoldDetail();
			frm._ID = ID;
			frm._isNew = false;
			frm.CongDoan = cboCongDoan.Text;
			frm.Updateby = txtCreateby.Text.Trim();
			if (frm.ShowDialog() == DialogResult.OK)
			{
				loadAllTextBox();
				if (frm._isNewTextBox)
				{
					deleteAllTextBox();
					creatNewTextBox(0, 0, "");
				}
			}
		}
		private void btnDelete_Click(object sender, EventArgs e)
		{
			if (grvKnifeJigWorking.SelectedRows == null) return;
			int selectedrowindex = grvKnifeJigWorking.SelectedCells[0].RowIndex;
			int ID = TextUtilsHP.ToInt(grvKnifeJigWorking.Rows[selectedrowindex].Cells["colID"].Value);
			string name = TextUtilsHP.ToString(grvKnifeJigWorking.Rows[selectedrowindex].Cells["colWorkingName"].Value);
			DialogResult result = MessageBox.Show("Bạn có muốn xóa mục kiểm tra [" + name + "] này không?", "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
			if (result == DialogResult.No) return;
			TextUtilsHP.ExcuteSQL($"Delete from dbo.KnifeJigWorking where ID = {ID}");
			loadAllTextBox();
		}
		private void toolStripButton2_Click(object sender, EventArgs e)
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

		private void grvKnifeJigWorking_Click(object sender, EventArgs e)
		{
			if (grvKnifeJigWorking.RowCount <= 0) return;
			int selectedrowindex = grvKnifeJigWorking.SelectedCells[0].RowIndex;
			lblx.Text = TextUtilsHP.ToString(grvKnifeJigWorking.Rows[selectedrowindex].Cells["colX"].Value);
			lbly.Text = TextUtilsHP.ToString(grvKnifeJigWorking.Rows[selectedrowindex].Cells["colY"].Value);

		}

		private void cboCongDoan_SelectedIndexChanged(object sender, EventArgs e)
		{
			loadAllTextBox();
			isZoom = 1;
			grvKnifeJigWorking.Visible = true;
			pdfViewer1.Visible = false;
			lblx.Visible = false;
			lbly.Visible = false;
		}

		private void btnZoom_Click(object sender, EventArgs e)
		{
			if (isZoom == 0)
			{
				isZoom = 1;
				grvKnifeJigWorking.Visible = true;
				pdfViewer1.Visible = false;
				lblx.Visible = false;
				lbly.Visible = false;
			}
			else
			{
				isZoom = 0;
				grvKnifeJigWorking.Visible = false;
				pdfViewer1.Visible = true;
				lblx.Visible = true;
				lbly.Visible = true;
			}
		}

		private void txtCreateby_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode != Keys.Enter) return;
			txtMaHang.Focus();
		}
	}
}
