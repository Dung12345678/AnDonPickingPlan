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
	public partial class frmHypMoldLine : _Forms
	{
		#region Variables
		int _KnifeDetailListID;
		DataTable _dt;
		DataTable _dtDisplay;
		DataTable _dtNoDisplay;
		string _path;
		Color _colorEmpty;
		CheckHistoryModel _checkHistory = new CheckHistoryModel();
		CheckHistoryDetailModel _checkHistoryDetailModel;
		#endregion Variables
		public frmHypMoldLine()
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
			txtDateTime.Text = DateTime.Now.ToString();
			SetInterface(false);
			Setbutton(false);
			loadPath();
			_colorEmpty = Color.FromArgb(255, 192, 255);
			initBackColor();

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
			btnSave.Enabled = isEdit;
		}
		private void SetInterface(bool isEdit)
		{
			btnSave.Visible = isEdit;
		}
		/// <summary>
		/// Sự kiện KeyDown của textbox mã hàng
		/// </summary>
		private void txtMaHang_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode != Keys.Enter) return;
			loadAllTextBox();
			grvKnifeJigWorking_CellClick(null, null);
			if (grvKnifeJigWorking.Rows.Count > 0)
			{
				grvKnifeJigWorking.CurrentCell = grvKnifeJigWorking[5, 0];
			}

			SetInterface(true);
		}
		private void loadAllTextBox()
		{
			try
			{
				//TODO: Xóa tất cả các textbox cũ
				deleteAllTextBox();
				//TODO: Load file PDF
				try
				{
					string path = _path + $"\\{txtMaHang.Text.Trim()}" + ".pdf";
					pdfViewer1.LoadDocument(path);
				}
				catch
				{
					MessageBox.Show("Không tìm thấy file PDF", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				}
				//TODO: Load Bảng mục kiểm tra theo ID của tên mã hàng
				DataTable dt = TextUtilsHP.Select($"select ID from dbo.KnifeDetailList where KnifeCode = '{txtMaHang.Text}'");
				if (dt.Rows.Count <= 0)
				{
					MessageBox.Show("Không tìm thấy mã khuôn", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
					Setbutton(false);
					return;
				}
				_KnifeDetailListID = TextUtilsHP.ToInt(dt.Rows[0]["ID"]);
				//List<int> ids = new List<int>(loadDataHypMold().Rows.Count);
				//foreach (DataRow row in loadDataHypMold().Rows) ids.Add((int)row["ID"]);

				//List<int> isDisplay = new List<int>(loadDataHypMold().Rows.Count);
				//foreach (DataRow row in loadDataHypMold().Rows) isDisplay.Add((int)row["isDisplay"]);
				//Show textbox

				//TODO: Lấy bảng Hiển thị text box từ bảng tổng
				DataRow[] dtr = loadDataHypMold().Select("isDisplay = 1");
				_dtDisplay = loadDataHypMold().Clone();
				_dtDisplay.Clear();
				foreach (DataRow dr in dtr)
				{
					_dtDisplay.ImportRow(dr);
				}
				_maxTextBoxNumber = _dtDisplay.Rows.Count - 1;
				//TODO: Show textbox
				for (int i = 0; i < _dtDisplay.Rows.Count; i++)
				{
					int LocationX = TextUtilsHP.ToInt(_dtDisplay.Rows[i]["X"]);
					int LocationY = TextUtilsHP.ToInt(_dtDisplay.Rows[i]["Y"]);
					//string RealValue = TextUtilsHP.ToString(loadDataHypMold().Rows[i]["RealValue"]);
					showTextBox(LocationX, LocationY, "", i);
				}
				Setbutton(true);
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
			string sql = $"Select * from dbo.KnifeJigWorking where KnifeDetailListID ={_KnifeDetailListID} order by SortOrder ";
			DataTable dt = TextUtilsHP.Select(sql);
			//Cách lấy data từ 1 bảng khác theo điều kiện
			//TODO: Lấy bảng không hiển thị textbox từ bảng tổng
			DataRow[] dtr = dt.Select("isDisplay = 0");
			_dtNoDisplay = dt.Clone();
			_dtNoDisplay.Clear();
			foreach (DataRow dr in dtr)
			{
				_dtNoDisplay.ImportRow(dr);
			}
			grvKnifeJigWorking.AutoGenerateColumns = false; // Không show tất cả cột lên
			grvKnifeJigWorking.DataSource = _dtNoDisplay;
			return dt;
		}
		private void showTextBox(int X, int Y, string RealValue, int row)
		{
			//TODO: Tạo textbox
			TextBox txtRow = new TextBox();
			txtRow.Name = "txtxxx" + $"{row}";
			txtRow.Font = new Font("Calibri", 10);
			txtRow.Size = new Size(50, 10);
			txtRow.Location = new Point(X, Y);
			txtRow.Dock = DockStyle.None;
			txtRow.BackColor = Color.FromArgb(255, 192, 255);
			txtRow.Enabled = true;
			txtRow.BringToFront();
			txtRow.Text = RealValue;
			panel2.Controls.Add(txtRow);
			pdfViewer1.SendToBack();
			txtRow.TextChanged += new System.EventHandler(this.Textbox_TextChanged);
			txtRow.KeyDown += new KeyEventHandler(Textbox_KeyDown);
			txtRow.Click += new EventHandler(Textbox_Click);
		}
		void initBackColor()
		{
			if (string.IsNullOrWhiteSpace(txtOrder.Text))
			{
				txtOrder.BackColor = _colorEmpty;
			}
			else
			{
				txtOrder.BackColor = Color.White;
			}
			if (string.IsNullOrWhiteSpace(txtMaHang.Text))
			{
				txtMaHang.BackColor = _colorEmpty;
			}
			else
			{
				txtMaHang.BackColor = Color.White;
			}
			if (string.IsNullOrWhiteSpace(txtWorkerCode.Text))
			{
				txtWorkerCode.BackColor = _colorEmpty;
			}
			else
			{
				txtWorkerCode.BackColor = Color.White;
			}
		}
		private void Textbox_TextChanged(object sender, EventArgs e)
		{
			TextBox txt = (TextBox)sender;
			if (string.IsNullOrWhiteSpace(txt.Text.Trim()))
			{
				txt.BackColor = _colorEmpty;
			}
			else
			{
				txt.BackColor = Color.White;
			}
		}
		private void Textbox_Click(object sender, EventArgs e)
		{
			TextBox txt = (TextBox)sender;
			//int row = TextUtilsHP.ToInt(txt.Name.Replace("txtxxx", ""));
			txt.SelectAll();
			//grvKnifeJigWorking.CurrentCell = grvKnifeJigWorking[5, row];
		}
		int _maxTextBoxNumber;
		private void Textbox_KeyDown(object sender, KeyEventArgs e)
		{
			TextBox txt = (TextBox)sender;
			int row = TextUtilsHP.ToInt(txt.Name.Replace("txtxxx", ""));
			//Chuyển lên text box trước nó nếu không có sẽ chuyển lên grv
			if (e.KeyCode == Keys.Up)
			{

				int RowPrevious = row - 1;
				if (RowPrevious < 0)
				{
					//Focus vào gridview
					grvKnifeJigWorking.Focus();
					grvKnifeJigWorking.CurrentCell = grvKnifeJigWorking[5, grvKnifeJigWorking.RowCount - 1];
					//grvKnifeJigWorking.FocusedColumn();
				}
				else
				{
					string nameControl = "txtxxx" + $"{row - 1}";
					Control tbox = this.Controls.Find(nameControl, true)[0];
					if (tbox != null)
					{
						(tbox as TextBox).Focus();
						(tbox as TextBox).SelectAll();
					}
				}


			}
			//Next textbox tiếp theo
			if (e.KeyCode != Keys.Enter) return;
			//TODO: Xử lý sự kiện keydown của textbox
			//TODO: So sánh giá trị 

			int MaxValue = TextUtilsHP.ToInt(_dtDisplay.Rows[row]["MaxValue"]);
			int MinValue = TextUtilsHP.ToInt(_dtDisplay.Rows[row]["MinValue"]);
			string StandartValue = TextUtilsHP.ToString(_dtDisplay.Rows[row]["PeriodValue"]);
			int CheckValueType = TextUtilsHP.ToInt(_dtDisplay.Rows[row]["CheckValueType"]);
			var value = txt.Text.Trim();
			if (CheckValueType == 2)
			{
				if (TextUtilsHP.ToString(value) != StandartValue)
				{
					txt.BackColor = Color.Red;
					txt.ForeColor = Color.White;
				}
				else
				{
					txt.BackColor = Color.Lime;
					txt.ForeColor = Color.Navy;
				}
			}
			else
			{
				if (TextUtilsHP.ToDecimal(value) != 0 && TextUtilsHP.ToDecimal(value) >= MinValue && TextUtilsHP.ToDecimal(value) <= MaxValue)
				{
					txt.BackColor = Color.Lime;
					txt.ForeColor = Color.Navy;
				}
				else
				{
					txt.BackColor = Color.Red;
					txt.ForeColor = Color.White;
				}
			}

			//TODO: Tìm textbox tiếp theo
			if (row <= _maxTextBoxNumber - 1)
			{
				string nameControl = "txtxxx" + $"{row + 1}";
				Control tbox = this.Controls.Find(nameControl, true)[0];
				if (tbox != null)
				{
					//foreach (Control item in panel2.Controls.OfType<TextBox>().ToList())
					//{
					//	if (item.Text == "")
					//	{
					//		item.BackColor = Color.White;
					//	}
					//}
					//(tbox as TextBox).ForeColor = Color.Navy;
					//(tbox as TextBox).BackColor = Color.Aqua;
					(tbox as TextBox).Focus();
					(tbox as TextBox).SelectAll();
				}
			}
			else
			{
				btnSave_Click(null, null);
			}
		}
		private void btnSave_Click(object sender, EventArgs e)
		{
			DialogResult result = MessageBox.Show("Bạn có chắc muốn cất dữ liệu?", "Cất?", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
			if (result == DialogResult.No) return;
			saveAtTable();
			//loadAllTextBox();
		}
		private void saveAtTable()
		{
			try
			{
				_checkHistory.KnifeDetailListID = _KnifeDetailListID;
				_checkHistory.OrderCode = txtOrder.Text.Trim();
				_checkHistory.CreateDate = DateTime.Now;
				_checkHistory.Type = 2;
				_checkHistory.KnifeDetailCode = txtMaHang.Text.Trim();
				_checkHistory.Machine = txtMachine.Text.Trim();
				_checkHistory.StepCode = cboCongDoan.Text.Trim();
				_checkHistory.WorkerCode = txtWorkerCode.Text.Trim();
				_checkHistory.ID = TextUtilsHP.ToInt(CheckHistoryBO.Instance.Insert(_checkHistory));
				List<CheckHistoryDetailModel> lst = new List<CheckHistoryDetailModel>();
				// Save CheckHistoryDetail
				for (int i = 0; i < grvKnifeJigWorking.Rows.Count; i++)
				{
					_checkHistoryDetailModel = new CheckHistoryDetailModel();
					_checkHistoryDetailModel.CheckHistoryID = _checkHistory.ID;
					_checkHistoryDetailModel.WorkingName = TextUtilsHP.ToString(grvKnifeJigWorking.Rows[i].Cells["colWorkingName"].Value);
					_checkHistoryDetailModel.CheckValueType = TextUtilsHP.ToInt(grvKnifeJigWorking.Rows[i].Cells["colCheckValueType"].Value);
					_checkHistoryDetailModel.Unit = TextUtilsHP.ToString(grvKnifeJigWorking.Rows[i].Cells["colUnit"].Value);
					_checkHistoryDetailModel.StandardValue = TextUtilsHP.ToString(grvKnifeJigWorking.Rows[i].Cells["colPeriodValue"].Value);
					_checkHistoryDetailModel.RealValue = TextUtilsHP.ToString(grvKnifeJigWorking.Rows[i].Cells["colRealvalue"].Value);
					_checkHistoryDetailModel.SortOrder = TextUtilsHP.ToInt(grvKnifeJigWorking.Rows[i].Cells["colSTT"].Value);
					_checkHistoryDetailModel.Result = TextUtilsHP.ToString(grvKnifeJigWorking.Rows[i].Cells["colEvaluate"].Value);
					_checkHistoryDetailModel.CreateDate = DateTime.Now;
					lst.Add(_checkHistoryDetailModel);

				}
				//TODO: Save at Textbox
				foreach (Control item in panel2.Controls.OfType<TextBox>().ToList())
				{
					_checkHistoryDetailModel = new CheckHistoryDetailModel();
					int row = TextUtilsHP.ToInt(item.Name.Replace("txtxxx", ""));
					string Result;
					if (item.BackColor == Color.Red)
					{
						Result = "NG";
					}
					else if (item.BackColor == Color.Lime)
					{
						Result = "OK";
					}
					else
					{
						Result = null;
					}
					_checkHistoryDetailModel.CheckHistoryID = _checkHistory.ID;
					_checkHistoryDetailModel.WorkingName = TextUtilsHP.ToString(_dtDisplay.Rows[row]["WorkingName"]);
					_checkHistoryDetailModel.CheckValueType = TextUtilsHP.ToInt(_dtDisplay.Rows[row]["CheckValueType"]);
					_checkHistoryDetailModel.Unit = TextUtilsHP.ToString(_dtDisplay.Rows[row]["Unit"]);
					_checkHistoryDetailModel.StandardValue = TextUtilsHP.ToString(_dtDisplay.Rows[row]["PeriodValue"]);
					_checkHistoryDetailModel.RealValue = TextUtilsHP.ToString(item.Text.Trim());
					_checkHistoryDetailModel.SortOrder = TextUtilsHP.ToInt(_dtDisplay.Rows[row]["SortOrder"]);
					_checkHistoryDetailModel.Result = Result;
					_checkHistoryDetailModel.CreateDate = DateTime.Now;
					lst.Add(_checkHistoryDetailModel);
				}
				Save(lst);
				grvKnifeJigWorking.DataSource = null;
				txtOrder.Text = "";
				txtMaHang.Text = "";
				cboCongDoan.Text = "";
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.ToString());
			}
		}
		async void Save(List<CheckHistoryDetailModel> checkHistoryDetailModel)
		{
			Task task = Task.Factory.StartNew(() =>
			{
				foreach (CheckHistoryDetailModel item in checkHistoryDetailModel)
				{
					CheckHistoryDetailBO.Instance.Insert(item);
				}
			});
			await task;

		}
		private void setColorTable(int condition)
		{
			if (condition == 1)
			{
				int selectedrowindex = grvKnifeJigWorking.SelectedCells[0].RowIndex;

				if (grvKnifeJigWorking.Rows[selectedrowindex].Cells["colEvaluate"].Value.ToString().Contains("OK"))
				{
					grvKnifeJigWorking.Rows[selectedrowindex].Cells["colEvaluate"].Style.BackColor = Color.Lime;
				}
				else if (grvKnifeJigWorking.Rows[selectedrowindex].Cells["colEvaluate"].Value.ToString().Contains("NG"))
				{
					grvKnifeJigWorking.Rows[selectedrowindex].Cells["colEvaluate"].Style.BackColor = Color.Red;
				}
			}

		}
		private void grvKnifeJigWorking_CellValueChanged(object sender, DataGridViewCellEventArgs e)
		{
			int selectedrowindex = 0;
			//	if (grvKnifeJigWorking.Rows.Count > 0)
			//	{
			//		selectedrowindex = grvKnifeJigWorking.SelectedCells[0].RowIndex;
			//	}
			//	else
			//	{
			//		string nameControl = "txtxxx0";
			//		Control tbox = this.Controls.Find(nameControl, true)[0];
			//		if (tbox != null)
			//		{
			//			//(tbox as TextBox).ForeColor = Color.Navy;
			//			//(tbox as TextBox).BackColor = Color.Aqua;
			//			(tbox as TextBox).Focus();
			//			(tbox as TextBox).SelectAll();
			//		}
			//	}

			//TODO: So sánh điều kiện của grid
			try
			{
				if (grvKnifeJigWorking.Rows.Count <= 0) return;
				int MaxValue = TextUtilsHP.ToInt(grvKnifeJigWorking.Rows[e.RowIndex].Cells["colMaxValue"].Value);
				int MinValue = TextUtilsHP.ToInt(grvKnifeJigWorking.Rows[e.RowIndex].Cells["colMinValue"].Value);
				string StandartValue = TextUtilsHP.ToString(grvKnifeJigWorking.Rows[e.RowIndex].Cells["colPeriodValue"].Value);
				int CheckValueType = TextUtilsHP.ToInt(grvKnifeJigWorking.Rows[e.RowIndex].Cells["colCheckValueType"].Value);
				int ID = TextUtilsHP.ToInt(grvKnifeJigWorking.Rows[e.RowIndex].Cells["colID"].Value);
				var value = grvKnifeJigWorking.Rows[e.RowIndex].Cells["colRealValue"].Value;
				string Evaluate = "";
				if (CheckValueType == 2)
				{
					if (TextUtilsHP.ToString(value).Trim().ToUpper() != StandartValue.Trim().ToUpper())
					{
						Evaluate = "NG";
						grvKnifeJigWorking.Rows[e.RowIndex].Cells["colEvaluate"].Value = "NG";

					}
					else
					{
						Evaluate = "OK";
						grvKnifeJigWorking.Rows[e.RowIndex].Cells["colEvaluate"].Value = "OK";
					}
				}
				else
				{
					if (TextUtilsHP.ToDecimal(value) != 0 && TextUtilsHP.ToDecimal(value) >= MinValue && TextUtilsHP.ToDecimal(value) <= MaxValue)
					{
						Evaluate = "OK";
						grvKnifeJigWorking.Rows[e.RowIndex].Cells["colEvaluate"].Value = "OK";

					}
					else
					{
						Evaluate = "NG";
						grvKnifeJigWorking.Rows[e.RowIndex].Cells["colEvaluate"].Value = "NG";
					}
				}
				setColorTable(1);

			}
			catch (Exception)
			{


			}
			//TODO: Điều kiện focus vào textbox
			if (_dtNoDisplay.Rows.Count == grvKnifeJigWorking.Rows.Count)
			{
				selectedrowindex = grvKnifeJigWorking.SelectedCells[0].RowIndex;
				if (selectedrowindex < grvKnifeJigWorking.Rows.Count - 1)
				{
					_nextrow = selectedrowindex;
				}
				else
				{
					_nextrow = selectedrowindex + 1;
				}

				if (_nextrow >= grvKnifeJigWorking.Rows.Count)
				{
					string nameControl = "txtxxx0";
					Control tbox = this.Controls.Find(nameControl, true)[0];
					if (tbox != null)
					{
						//(tbox as TextBox).ForeColor = Color.Navy;
						//(tbox as TextBox).BackColor = Color.Aqua;
						(tbox as TextBox).Focus();
						(tbox as TextBox).SelectAll();
					}
				}
			}


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

		private void grvKnifeJigWorking_CellClick(object sender, DataGridViewCellEventArgs e)
		{
			grvKnifeJigWorking.Focus();
		}
		int _nextrow;
		private void grvKnifeJigWorking_KeyDown(object sender, KeyEventArgs e)
		{
			//if (e.KeyCode != Keys.Enter) return;
			//try
			//{
			//	int selectedrowindex = grvKnifeJigWorking.SelectedCells[0].RowIndex;
			//	//TODO: Điều kiện focus vào textbox
			//	if (_dtNoDisplay.Rows.Count == grvKnifeJigWorking.Rows.Count)
			//	{

			//		if (selectedrowindex < grvKnifeJigWorking.Rows.Count - 1)
			//		{
			//			_nextrow = selectedrowindex;
			//		}
			//		else
			//		{
			//			_nextrow = selectedrowindex + 1;
			//		}

			//		if (_nextrow >= grvKnifeJigWorking.Rows.Count)
			//		{
			//			string nameControl = "txtxxx0";
			//			Control tbox = this.Controls.Find(nameControl, true)[0];
			//			if (tbox != null)
			//			{
			//				(tbox as TextBox).ForeColor = Color.Navy;
			//				(tbox as TextBox).BackColor = Color.Aqua;
			//				(tbox as TextBox).Focus();
			//				(tbox as TextBox).SelectAll();
			//			}
			//		}
			//		//int isDisplay = TextUtilsHP.ToInt(grvKnifeJigWorking.Rows[selectedrowindex].Cells["colDisplay"].Value);
			//		//if (isDisplay == 1)
			//		//{
			//		//    grvKnifeJigWorking.Rows[selectedrowindex].Cells["colRealValue"].ReadOnly = true;
			//		//    string nameControl = "txtxxx" +$"{e.RowIndex}";
			//		//    Control tbox = this.Controls.Find(nameControl, true)[0];
			//		//    if (tbox!=null)
			//		//    {
			//		//        (tbox as TextBox).Focus();
			//		//        (tbox as TextBox).SelectAll();
			//		//    }
			//		//}
			//		//else
			//		//{
			//		//    grvKnifeJigWorking.Focus();
			//		//    //grvKnifeJigWorking.CurrentCell= grvKnifeJigWorking[grvKnifeJigWorking.CurrentCell.ColumnIndex + 1, grvKnifeJigWorking.CurrentCell.RowIndex];
			//		//    grvKnifeJigWorking.Rows[selectedrowindex].Cells["colRealValue"].ReadOnly = false;
			//		//}
			//	}

			//}
			//catch (Exception)
			//{


			//}
		}
	}
}
