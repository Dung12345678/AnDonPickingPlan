using BMS.Business;
using BMS.Model;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
using BMS.Utils;
using System.Collections;
using System.IO;
using System.Threading.Tasks;
using System.Threading;
using System.Linq;

namespace BMS
{
	public delegate void LapLai(DataGridViewRow currentRow);


	public partial class frmTestta2 : Form
	{


		public long _WorkerID = 0;
		private Color _mauHongNhat = Color.FromArgb(255, 204, 255);
		private Color _mauXam = Color.FromArgb(191, 191, 191);
		private Color _mauCarot = Color.FromArgb(244, 177, 131);
		private Color _mauTanSuat = Color.FromArgb(0, 176, 80);


		int _countGearWorking = 0;
		public string _Order;

		string _pathTester = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Tester");
		string _pathSeverImage = @"Testler\Image";
		string _pathSeverAudio = @"Testler\Audio";

		Dictionary<string, string> _dicLanguageControls;
		Dictionary<string, string> _dicLanguageGrv;
		string[] _arrLangGrv;
		string[] _arrLangLabel;
		string _language = "VN";


		public frmTestta2()
		{
			InitializeComponent();
		}



		private void frmTestta2_Load(object sender, EventArgs e)
		{
			DocUtils.InitFTPQLSX();

			colGear.Frozen = true;
			// load file language
			if (!File.Exists(Path.Combine(Application.StartupPath, "LanguageGrv.txt"))) return;
			if (!File.Exists(Path.Combine(Application.StartupPath, "Language.txt"))) return;
			_arrLangGrv = File.ReadAllLines(Path.Combine(Application.StartupPath, "LanguageGrv.txt"));
			_arrLangLabel = File.ReadAllLines(Path.Combine(Application.StartupPath, "Language.txt"));
			cboLanguage.SelectedIndex = 0;
			//loadFileLanguage(_language, _arrLangLabel, _arrLangGrv);
			LoadDataTestlerByOrder(_Order);


		}
		private void loadFileLanguage(string lang, string[] arrLangLabel, string[] arrLangGrv)
		{

			_dicLanguageControls = new Dictionary<string, string>();
			_dicLanguageGrv = new Dictionary<string, string>();
			for (int i = 0; i < arrLangGrv.Length; i++)
			{
				if (string.IsNullOrWhiteSpace(arrLangGrv[i])) continue;

				if (lang == "VN")
				{
					_dicLanguageGrv.Add(arrLangGrv[i].Split(';')[0].Trim(), arrLangGrv[i].Split(';')[1].Trim());
				}
				else
				{
					_dicLanguageGrv.Add(arrLangGrv[i].Split(';')[0].Trim(), arrLangGrv[i].Split(';')[2].Trim());
				}
			}

			for (int i = 0; i < arrLangLabel.Length; i++)
			{
				if (string.IsNullOrWhiteSpace(arrLangLabel[i])) continue;

				if (lang == "VN")
				{
					_dicLanguageControls.Add(arrLangLabel[i].Split(';')[0].Trim(), arrLangLabel[i].Split(';')[1].Trim());

				}
				else
				{
					_dicLanguageControls.Add(arrLangLabel[i].Split(';')[0].Trim(), arrLangLabel[i].Split(';')[2].Trim());

				}
			}
		}
		private void changeLanguage()
		{
			loadFileLanguage(_language, _arrLangLabel, _arrLangGrv);
			// thay đổi ngôn ngữ cho các control label
			foreach (var item in _dicLanguageControls)
			{
				Control[] ctr = this.Controls.Find(item.Key, true);
				if (ctr.Length > 0) ctr[0].Text = item.Value;
			}
			// thay đổi ngôn ngữ cho collumn trong grv
			foreach (var item in _dicLanguageGrv)
			{
				if (grvData.Columns[item.Key] == null) continue;

				grvData.Columns[item.Key].HeaderText = item.Value;
			}
	
		}
		int[] getFocusRowTarget(int rowIndex)
		{
			//int row = 4;
			int[] indexRowCol = new int[2];
			int countRow = grvData.Rows.Count;
			for (int i = rowIndex + 1; i < countRow; i++)
			{
				for (int j = 3; j <= (colResult1.Index - 1); j++)
				{
					//string value = TextUtils.ToString(grvData.Rows[i].Cells[j].Value);
					if (grvData.Rows[i].Cells[j].Style.BackColor != _mauXam)
					{
						indexRowCol[0] = i;
						indexRowCol[1] = j;
						return indexRowCol;
					}
				}
			}
			indexRowCol[0] = 4;
			indexRowCol[1] = 3;
			return indexRowCol;
			// return row;
		}
		void setFocusCell(int rowIndex)
		{
			try
			{
				int colCurrent = grvData.CurrentCell.ColumnIndex;

				for (int j = colCurrent + 1; j < colQtyLap.Index; j++)
				{
					//if (j >= colResult1.Index - 1) continue;                   
					//string value = TextUtils.ToString(grvData.Rows[rowIndex].Cells[j].Value);
					if (grvData[j, rowIndex].Style.BackColor != _mauXam)
					{
						string filename = getPathFile(grvData.CurrentCell);
						string nameworking = grvData.CurrentCell.OwningColumn.HeaderText;

						grvData.CurrentCell = grvData[j, rowIndex];                       
				
						if (j == colResult1.Index)
						{
							grvData.CurrentCell = grvData[grvData.CurrentCell.ColumnIndex + _qtyAudio + 1, rowIndex];
							grvData.CurrentCell.ReadOnly = false;
						
						}

						return;
					}
				}
				int[] rowIndex1 = getFocusRowTarget(rowIndex);
				for (int i = 1; i <= rowIndex1[0] - rowIndex; i++)
				{
					_isCatchKey = false;
					SendKeys.Send("{DOWN}");
				}
				grvData.CurrentCell = grvData[rowIndex1[1], rowIndex1[0]];
				grvData.CurrentCell.ReadOnly = false;

			}
			catch (Exception)
			{
			}
		}
		bool _isCatchKey = false;
		private void grvData_CellEndEdit(object sender, DataGridViewCellEventArgs e)
		{
			if (e.RowIndex < 4) return;   //Bỏ 4 dòng đầu đi không nhảy nhót gì cả

			if (e.ColumnIndex == colResult1.Index) return;

			if (e.RowIndex < grvData.Rows.Count - 1)
			{
				_isCatchKey = false;
				SendKeys.Send("{Up}");

			}

			if (e.ColumnIndex >= 1 && e.ColumnIndex < colQtyLap.Index)
			{
                               
                //grvData.CurrentCell.Value = grvData.CurrentCell.Valu
                //Xong thì nhảy đến cột đánh giá tiếp theo
                setFocusCell(e.RowIndex);
                
            }
		}


		private void LoadDataTestlerByOrder(string order)
		{
			// load data từ store
			DataSet dts = TextUtils.GetListDataFromSP("spGetResultData_ByOrderCode"
					, new string[1] { "@OrderCode" }
					, new object[1] { order });
			DataTable dataTable0 = new DataTable();
			DataTable dataTable1 = new DataTable();
			if (dts.Tables.Count > 1)
			{
				dataTable0 = dts.Tables[0];
				dataTable1 = dts.Tables[1];
			}

			// kieeemr tra table 0 
			if (dataTable0.Rows.Count > 0)
			{
				//nếu có thì binding dữ liệu vào các control dạng text
				txtOrderNo.Text = _Order;
				txtWorkerName.Text = TextUtils.ToString(dataTable0.Rows[0]["WorkerCode"]);
				txtHYP.Text = TextUtils.ToString(dataTable0.Rows[0]["HypCode"]);
				txtQty.Text = TextUtils.ToString(dataTable0.Rows[0]["Qty"]);
				dteNgayGiaCong.Value = TextUtils.ToDate3(dataTable0.Rows[0]["DateLR"]);
				txtBatch.Text = TextUtils.ToString(dataTable0.Rows[0]["Batch"]);
				txtConfirmer.Text = TextUtils.ToString(dataTable0.Rows[0]["Confirmer"]);
				lblTester.Text = TextUtils.ToString(dataTable0.Rows[0]["TesterName"]);
				_sttStart = (TextUtils.ToInt(dataTable0.Rows[0]["SttStart"]));
				_hypCode = txtHYP.Text.Trim();
				loadGear(_hypCode);
				GenerateByQty();

				//binding table 1 vào grid
				if (dataTable1.Rows.Count > 0)
				{
					bindingDataRow(dataTable1);
					checkAudioImageExist();

				}
			}
			else
			{
				// focus khi không load đc order trong db
				txtConfirmer.Focus();
			}
		}

		private void txtOrderNo_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				LoadDataTestlerByOrder(txtOrderNo.Text.Trim());
			}
		}

        void checkAudioImageExist()
        {
            for (int i = 0; i < grvData.Rows.Count; i++)
            {
                if (i < 4 || grvData.Rows[i].Cells[colResult1.Index].Style.BackColor == _mauXam) continue;
                string stt = TextUtils.ToString(grvData.Rows[i].Cells[colNumber.Index].Value);
                string order = TextUtils.ToString(txtOrderNo.Text.Trim());
                string qtyLap = TextUtils.ToString(grvData.Rows[i].Cells[colQtyLap.Index].Value);
                for (int h = 1; h <= _qtyImage; h++)
                {
                    string vtriAnh = TextUtils.ToString(grvData.Columns["colImage" + h].Tag);
                    string filename = Path.Combine(_pathSeverImage, order + "_" + stt + "_" + qtyLap + "_" + vtriAnh + ".jpg");
                    if (DocUtils.CheckExits(filename))
                    {
                        grvData.Rows[i].Cells["colImage" + h].Value = "View";
                    }

                }

                for (int h = 1; h <= _qtyAudio; h++)
                {
                    string vtriAnh = TextUtils.ToString(grvData.Columns["colAudio" + h].Tag);
                    string filename = Path.Combine(_pathSeverAudio, order + "_" + stt + "_" + qtyLap + "_" + vtriAnh + ".wav");
					if (DocUtils.CheckExits(filename))
					{
						grvData.Rows[i].Cells["colAudio" + h].Value = "Play";
                    }

                }
            }
        }

		void bindingDataRow(DataTable dtGridData)
		{
			for (int i = 0; i < grvData.Rows.Count; i++)
			{
				int stt = TextUtils.ToInt(grvData.Rows[i].Cells[colNumber.Index].Value);
				if (i < 4) continue;
				DataRow[] arrRow = dtGridData.Select("STT = " + stt);
				if (arrRow.Length == 0) continue;
				for (int j = 0; j < grvData.Columns.Count; j++)
				{
					DataGridViewColumn col = grvData.Columns[j];
					if (col.DataPropertyName == "image" || col.DataPropertyName == "audio")
					{
                       
                        continue;
					}
					string valueRow = TextUtils.ToString(arrRow[0][col.DataPropertyName]);
					if (string.IsNullOrWhiteSpace(valueRow)) continue;
					grvData.Rows[i].Cells[j].Value = valueRow;

				}

			}
			//Set focus vào ô khi load lại dữ liệu
			grvData.Focus();
			setFocusByLoadOrder();
		}

		private void setFocusByLoadOrder()
		{
			bool _isFind = false;
			for (int i = 0; i < grvData.Rows.Count; i++)
			{
				if (i < 4) continue;
				for (int j = colGear.Index + 1; j < colResult1.Index; j++)
				{
					if (grvData.Rows[i].Cells[j].Style.BackColor == _mauHongNhat)
					{
						grvData.CurrentCell = grvData.Rows[i].Cells[j];
						grvData.CurrentCell.ReadOnly = false;
						_isFind = true;
						break;
					}
				}
				if (_isFind)
					break;
			}
		}
		private void GenerateByQty()
		{
			int qty = TextUtils.ToInt(txtQty.Text.Trim());
			if (qty == 0)
			{
				return;
			}
			// trừ đi 4 dòng đầu để stt bắt đầu từ 0
			if (qty == 5) _tuneParam = -4;
			else if (qty != 5 && qty >= 0)
			{
                _tuneParam = (_sttStart - 4);
            }
			


			if (!GenerateByHpy(qty)) return;


			//remove các dòng đã tạo nếu có
			if (grvData.Rows.Count > 4)
			{
				for (int i = grvData.Rows.Count - 1; i > 3; i--)
				{
					grvData.Rows.RemoveAt(i);
				}
			}
			if (qty == 5)
			{
				qty += 1;
			}
			grvData.Rows.Add(qty);
			//Thêm số thứ tự + Bôi màu cho các ô theo tần suất check
			for (int i = 0; i < grvData.Rows.Count; i++)
			{
				if (i < 4) continue;
				grvData.Rows[i].Cells[0].Value = i + _tuneParam;
				grvData.Rows[i].Cells[colQtyLap.Index].Value = 0;

				int countXam = 0;
				for (int j = 1; j < grvData.Columns.Count; j++)
				{
					// locked các cell lại. Các cell được mở khi các cell trước đó đã có giá trị.
					grvData.Rows[i].Cells[j].ReadOnly = true;
					string tagValue = TextUtils.ToString(grvData.Columns[j].Tag);
					//nếu là những cột tự sinh, là mục kiểm tra
					if (!string.IsNullOrEmpty(tagValue))
					{
						if (qty <= 6)
						{

							if (j <= colResult1.Index)
							{
								grvData.Rows[i].Cells[j].Style.BackColor = _mauHongNhat;

							}
						}
						else
						{
							string tansuat = TextUtils.ToString(grvData.Rows[3].Cells[j].Value);
							if (tansuat.Contains("/"))
							{
								string[] arr = tansuat.Split('/');
								int moc = TextUtils.ToInt(arr[1]);
								// nếu vị trí hàng chia hết cho tần suất. thực hiện đổi màu nền tại hàng đó là LightBlue
								if ((i - 4) % moc == 0)
								{
									grvData.Rows[i].Cells[j].Style.BackColor = _mauHongNhat;

								}
								else
								{
									grvData.Rows[i].Cells[j].Value = "N";
									grvData.Rows[i].Cells[j].Style.BackColor = _mauXam;
									grvData.Rows[i].Cells[j].Style.ForeColor = _mauXam;
									grvData.Rows[i].Cells[j].Selected = false;
									grvData.Rows[i].Cells[j].ReadOnly = true;
									countXam++;
								}
							}
						}

					}
					else
					{
						if (j == colHyp.Index || j == colGear.Index || j == colResult1.Index)
						{
							grvData.Rows[i].Cells[j].Style.BackColor = _mauHongNhat;

						}

					}

				}
				if (countXam == _countGearWorking)
				{
					grvData.Rows[i].Cells[1].Style.BackColor = _mauXam;
					grvData.Rows[i].Cells[1].Selected = false;
					grvData.Rows[i].Cells[1].ReadOnly = true;

					grvData.Rows[i].Cells[2].Style.BackColor = _mauXam;
					grvData.Rows[i].Cells[2].Selected = false;
					grvData.Rows[i].Cells[2].ReadOnly = true;

					grvData.Rows[i].Cells[colResult1.Index].Style.BackColor = _mauXam;
					grvData.Rows[i].Cells[colResult1.Index].Selected = false;
					grvData.Rows[i].Cells[colResult1.Index].ReadOnly = true;

					for (int h = 1; h <= _qtyImage; h++)
					{

						grvData.Rows[i].Cells["colImage" + h].Style.BackColor = _mauXam;
					}

					for (int h = 1; h <= _qtyAudio; h++)
					{
						grvData.Rows[i].Cells["colAudio" + h].Style.BackColor = _mauXam;
					}


				}
				else
				{
					
				}
			}


		}

		private bool GenerateByHpy(int qty)
		{
			if (_hypCode == "")
			{
				MessageBox.Show("HYP không được để trống");
				return false;
			}
			DataTable dtGearWorking = TextUtils.LoadDataFromSP("spGetGearWorking_ByHypCode", "A", new string[] { "@HYP" }, new object[] { _hypCode });
			Expression exp = new Expression("HYP", _hypCode);
			ArrayList arr = GearBO.Instance.FindByExpression(exp);


			// Từ các hạng mục kiểm tra sinh ra các cột
			_countGearWorking = dtGearWorking.Rows.Count;

			//Đầu tiên bỏ các dòng hiện tại đang có trên danh sách đi
			grvData.Rows.Clear();
			//Xóa các cột có tag Value
			List<int> lstColumnTag = new List<int>();
			for (int i = 0; i < grvData.Columns.Count; i++)
			{
				DataGridViewColumn col = grvData.Columns[i];
				if (col.Tag != null && col.Tag.ToString() != "")
				{
					lstColumnTag.Add(col.Index);
				}
			}
			for (int i = lstColumnTag.Count - 1; i >= 0; i--)
			{
				grvData.Columns.Remove(grvData.Columns[lstColumnTag[i]]);
			}
			//add tiếp 4 dòng
			for (int i = 0; i < 4; i++)
			{
				grvData.Rows.Add();
			}
			grvData.Rows[0].Cells[0].Value = "Tiêu chuẩn";
			grvData.Rows[1].Cells[0].Value = "MIN";
			grvData.Rows[2].Cells[0].Value = "MAX";
			grvData.Rows[3].Cells[0].Value = "Tần suất";


			if (_countGearWorking == 0 || arr.Count <= 0)
			{
				MessageBox.Show("Bạn phải nhập đúng HYP!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return false;
			}

			GearModel gear = arr[0] as GearModel;
			_qtyImage = gear.QtyImage;
			_qtyAudio = gear.QtyAudio;
			if (qty <= 5)
			{
				_qtyImage += 2;
			}
			for (int i = _countGearWorking - 1; i >= 0; i--)
			{
				DataGridViewTextBoxColumn col = new DataGridViewTextBoxColumn();
				col.HeaderText = TextUtils.ToString(dtGearWorking.Rows[i]["WorkingName"]);//Tên của GearWorking
				col.Tag = TextUtils.ToString(dtGearWorking.Rows[i]["ID"]) + ";" + TextUtils.ToString(dtGearWorking.Rows[i]["SortOrder"]);//Cái này để lưu lại ID của GearWorking
				col.DataPropertyName = TextUtils.ToString(dtGearWorking.Rows[i]["SortOrder"]);//cho stt vào
				col.Width = 70;
				col.SortMode = DataGridViewColumnSortMode.NotSortable;
				col.Name = "col" + i;
				grvData.Columns.Insert(3, col);
				// Cập nhật thông tin dữ liệu 4 dòng dữ liệu với giá trị tiêu chuẩn tương ứng với mỗi cột hạng mục
				grvData.Rows[0].Cells[3].Value = TextUtils.ToDecimal(dtGearWorking.Rows[i]["DefaultValue"]);
				grvData.Rows[1].Cells[3].Value = TextUtils.ToDecimal(dtGearWorking.Rows[i]["MinValue"]);
				grvData.Rows[2].Cells[3].Value = TextUtils.ToDecimal(dtGearWorking.Rows[i]["MaxValue"]);
				grvData.Rows[3].Cells[3].Value = TextUtils.ToString(dtGearWorking.Rows[i]["TanSuat"]);

				grvData.Rows[0].Cells[3].Style.BackColor = _mauCarot;
				grvData.Rows[1].Cells[3].Style.BackColor = _mauCarot;
				grvData.Rows[2].Cells[3].Style.BackColor = _mauCarot;
				grvData.Rows[3].Cells[3].Style.BackColor = _mauTanSuat;

				

			}

			for (int i = _qtyAudio; i >= 1; i--)
			{
				DataGridViewTextBoxColumn col = new DataGridViewTextBoxColumn();
				col.HeaderText = "Âm thanh " + i;//Tên của GearWorking
												 //col.Tag = TextUtils.ToString(dtGearWorking.Rows[i]["ID"]) + ";" + TextUtils.ToString(dtGearWorking.Rows[i]["SortOrder"]);//Cái này để lưu lại ID của GearWorking
												 //col.DataPropertyName = i.ToString();//cho stt vào
				col.DataPropertyName = "audio";
				col.Tag = i;
				col.Width = 50;
				col.SortMode = DataGridViewColumnSortMode.NotSortable;
				col.Name = "colAudio" + i;
				col.DefaultCellStyle.ForeColor = Color.Blue;
				grvData.Columns.Insert(colResult1.Index + 1, col);
			}

			for (int i = _qtyImage; i >= 1; i--)
			{
				DataGridViewTextBoxColumn col = new DataGridViewTextBoxColumn();
				col.HeaderText = "Ảnh " + i;//Tên của GearWorking
											//col.Tag = TextUtils.ToString(dtGearWorking.Rows[i]["ID"]) + ";" + TextUtils.ToString(dtGearWorking.Rows[i]["SortOrder"]);//Cái này để lưu lại ID của GearWorking
				col.DataPropertyName = "image";//cho stt vào
				col.Tag = i;
				col.Width = 50;
				col.SortMode = DataGridViewColumnSortMode.NotSortable;
				col.Name = "colImage" + i;
				col.DefaultCellStyle.ForeColor = Color.Blue;
				grvData.Columns.Insert(colResult1.Index + 1 + _qtyAudio, col);

			}

			grvData.Rows[0].ReadOnly = true;
			grvData.Rows[1].ReadOnly = true;
			grvData.Rows[2].ReadOnly = true;
			grvData.Rows[3].ReadOnly = true;

			return true;

		}


		private void txt_TextChanged(object sender, EventArgs e)
		{
			TextBox txt = (TextBox)sender;
			if (string.IsNullOrWhiteSpace(txt.Text.Trim()))
			{
				txt.BackColor = _mauHongNhat;
			}
			else
			{
				txt.BackColor = Color.White;
			}
		}
		/// <summary>
		/// F12 is Save
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		/// 


		private bool ValidateData()
		{
			if (txtConfirmer.Text.Trim() == "")
			{
				MessageBox.Show("Bạn chưa nhập người xác nhận", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return false;
			}
			if (txtBatch.Text.Trim() == "")
			{
				MessageBox.Show("Bạn chưa nhập Lô lot", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return false;
			}
			if (txtOrderNo.Text.Trim() == "")
			{
				MessageBox.Show("Bạn chưa nhập Order", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return false;
			}

			if (txtHYP.Text.Trim() == "")
			{
				MessageBox.Show("Bạn chưa nhập HYP", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return false;
			}
			if (cboGear.SelectedIndex == -1)
			{
				MessageBox.Show("Bạn chưa nhập Gear", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return false;
			}
			if (TextUtils.ToInt(txtQty.Text.Trim()) == 0)
			{
				MessageBox.Show("Bạn chưa nhập Số lượng", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return false;
			}
			if (txtWorkerName.Text.Trim() == "")
			{
				MessageBox.Show("Bạn chưa nhập Người gia công", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return false;
			}
			return true;
		}

		/// <summary>
		/// Sự kiện khi next ô tiếp theo khi đã điền dữ liệu vào ô trên grid
		/// created by: nguyenvan.thao
		/// created date: 25.12.2019
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>

		private void grvData_CellValueChanged(object sender, DataGridViewCellEventArgs e)
		{
			if (e.RowIndex < 4) return;   //Bỏ 4 dòng đầu đi không nhảy nhót gì cả
										  //if (e.ColumnIndex == colCheckEye11.Index || e.ColumnIndex == colCheckEye12.Index || e.ColumnIndex == colQtyLap.Index) return;

			string value = TextUtils.ToString(grvData.Rows[e.RowIndex].Cells[e.ColumnIndex].Value);

			if (string.IsNullOrWhiteSpace(value))
			{
				grvData.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = _mauHongNhat;
			}
			else
			{
				if (value == "N") return;
				if (e.ColumnIndex == colHyp.Index)
				{
					if (!value.ToUpper().Contains(txtHYP.Text.ToUpper().Trim()))
					{
						grvData.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = Color.Red;
						grvData.Rows[e.RowIndex].Cells[colResult1.Index].Style.BackColor = Color.Red;
						grvData.Rows[e.RowIndex].Cells[colResult1.Index].Value = "NG";
					}
					else
					{
						grvData.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = Color.White;
					}
				}
				if (e.ColumnIndex == colGear.Index)
				{
					if (!value.ToUpper().Contains(cboGear.Text.ToUpper().Trim()))
					{
						grvData.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = Color.Red;
						grvData.Rows[e.RowIndex].Cells[colResult1.Index].Style.BackColor = Color.Red;
						grvData.Rows[e.RowIndex].Cells[colResult1.Index].Value = "NG";
					}
					else
					{
						grvData.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = Color.White;
					}
				}
				if (e.ColumnIndex > 2 && e.ColumnIndex < colResult1.Index)
				{
					decimal min = TextUtils.ToDecimal(grvData.Rows[1].Cells[e.ColumnIndex].Value);
					decimal max = TextUtils.ToDecimal(grvData.Rows[2].Cells[e.ColumnIndex].Value);
					decimal currentValue = TextUtils.ToDecimal(value);
					if (currentValue > max || currentValue < min)
					{
						grvData.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = Color.Red;
						grvData.Rows[e.RowIndex].Cells[colResult1.Index].Style.BackColor = Color.Red;
						grvData.Rows[e.RowIndex].Cells[colResult1.Index].Value = "NG";
					}
					else
					{
						grvData.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = Color.White;

					}
				}
				if (e.ColumnIndex == colResult1.Index)
				{
					if (value == "OK")
					{
						grvData.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = Color.White;
					}
					else
					{
						grvData.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = Color.Red;
					}
				}
			}
			// gọi hàm check ok/ng, đổi màu ô kết quả
			checkOKNG(e.RowIndex, e.ColumnIndex);

		}

		private void checkOKNG(int rowIndex, int colIndex)
		{
			if (rowIndex < 4 || colIndex < 3 || colIndex >= colResult1.Index) return;
			int countColorRed = 0;
			int countColorPink = 0;
			int countColorWhite = 0;
			for (int i = colGear.Index + 1; i < colResult1.Index - 1; i++)
			{
				if (grvData.Rows[rowIndex].Cells[i].Style.BackColor == Color.Red) countColorRed++;
				if (grvData.Rows[rowIndex].Cells[i].Style.BackColor == _mauHongNhat) countColorPink++;
				if (grvData.Rows[rowIndex].Cells[i].Style.BackColor == Color.White) countColorWhite++;
			}
			if (countColorRed > 0)
			{
				grvData.Rows[rowIndex].Cells[colResult1.Index].Style.BackColor = Color.Red;
				grvData.Rows[rowIndex].Cells[colResult1.Index].Value = "NG";
                return;
			}
			if (countColorPink > 0)
			{
				grvData.Rows[rowIndex].Cells[colResult1.Index].Style.BackColor = _mauHongNhat;
				grvData.Rows[rowIndex].Cells[colResult1.Index].Value = "";
				return;
			}
			grvData.Rows[rowIndex].Cells[colResult1.Index].Style.BackColor = Color.White;
			grvData.Rows[rowIndex].Cells[colResult1.Index].Value = "OK";

		}

		/// <summary>
		/// Sự kiện dùng để xem lại ảnh khi muốn check bằng mắt thường
		/// created by: nvthao
		/// created date: 25.12.2019
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		/// 

		private void grvData_CellClick(object sender, DataGridViewCellEventArgs e)
		{
						
			try
			{
				string filename = getPathFile(grvData.CurrentCell);
				if (filename == "") return;
				if (e.ColumnIndex > colResult1.Index + _qtyAudio && e.ColumnIndex < colQtyLap.Index)
				{
					string pfm = Path.Combine(_pathTester, filename + ".jpg");
					if (File.Exists(pfm))
					{
						Process.Start(pfm);
					}
					else
					{
						string pathFileServer = Path.Combine(_pathSeverImage, filename + ".jpg");
						if (DocUtils.CheckExits(pathFileServer))
						{
							DocUtils.DownloadFile(_pathTester, filename + ".jpg", pathFileServer);
							Process.Start(pfm);
						}

					}
				}
				else if (e.ColumnIndex > (colResult1.Index) && e.ColumnIndex <= colResult1.Index + _qtyAudio)
				{
					string pfo = "";

					pfo = Path.Combine(_pathTester, filename+".wav");
					if (File.Exists(pfo))
					{
						Process.Start(pfo);
					}
					else
					{
						string pathFileServer = Path.Combine(_pathSeverAudio, filename+".wav");
						if (DocUtils.CheckExits(pathFileServer))
						{
							DocUtils.DownloadFile(_pathTester, filename + ".wav", pathFileServer);
							Process.Start(pfo);
						}
					}
				}

			}
			catch
			{

			}
		}


		private void button1_Click(object sender, EventArgs e)
		{
			DataGridViewTextBoxColumn col = new DataGridViewTextBoxColumn();
			col.HeaderText = "Hi there";//Tên của GearWorking
			col.Tag = 0;//Cái này để lưu lại ID của GearWorking
			col.DataPropertyName = "";//cho stt vào
			col.Width = 90;
			col.SortMode = DataGridViewColumnSortMode.NotSortable;
			grvData.Columns.Insert(3, col);
		}

		ArrayList _arrGear;
		int _qtyImage;
		int _qtyAudio;
		string _hypCode = "";
		private void txtHYP_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode != Keys.Enter) return;
			// Lẩy ra danh sách hạng mục cần kiểm tra

			grvData.Rows.Clear();

			_hypCode = txtHYP.Text.Trim();
			if (_hypCode == "")
			{
				MessageBox.Show("HYP không được để trống");
				return;
			}
			loadGear(_hypCode);
			txtQty.Focus();

		}

		private void loadGear(string hypCode)
		{
			Expression exp = new Expression("HYP", hypCode);
			ArrayList arr = GearBO.Instance.FindByExpression(exp);
			if (arr.Count > 0)
			{
				GearModel gear = arr[0] as GearModel;
				_arrGear = new ArrayList();
				_arrGear.Add(TextUtils.ToString(gear.Gear1));
				_arrGear.Add(TextUtils.ToString(gear.Gear2));
				_arrGear.Add(TextUtils.ToString(gear.Gear3));

				if (_arrGear.Count > 0)
				{
					cboGear.DataSource = _arrGear;
					cboGear.SelectedIndex = 0;
				}
				else
				{
					cboGear.SelectedIndex = -1;
				}
			}
			

		}

		int _tuneParam = 0;
		int _sttStart = 0;
		/// <summary>
		/// Sinh số lượng sản phẩm cần kiểm tra
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void txtQty_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode != Keys.Enter) return;


			int qty = TextUtils.ToInt(txtQty.Text.Trim());
			if (qty == 0)
			{
				return;
			}
			// trừ đi 4 dòng đầu để stt bắt đầu từ 0
			if (qty == 5) _tuneParam = -4;
			else if(qty != 5 && qty >= 0)
			{
                frmStt frm = new frmStt();
                frm.ShowDialog();
                if (frm.DialogResult == DialogResult.OK)
                {
                    _sttStart = frm._SttStart;
                    _tuneParam = (_sttStart - 4);
                }
                else
                {
                    _tuneParam = -4;
                }
            }
			
			
			if (!GenerateByHpy(qty)) return;


			//remove các dòng đã tạo nếu có
			if (grvData.Rows.Count > 4)
			{
				for (int i = grvData.Rows.Count - 1; i > 3; i--)
				{
					grvData.Rows.RemoveAt(i);
				}
			}

			// Tạo ra tương ứng các dòng trên grid bằng giá trị số lượng của ô nhập
			//grvData.Rows.Add(qty);
			if (qty == 5)
			{
				qty += 1;
			}
			grvData.Rows.Add(qty);

			//Thêm số thứ tự + Bôi màu cho các ô theo tần suất check
			for (int i = 0; i < grvData.Rows.Count; i++)
			{
				if (i < 4) continue;
				grvData.Rows[i].Cells[0].Value = i + _tuneParam;
				grvData.Rows[i].Cells[colQtyLap.Index].Value = 0;

				int countXam = 0;
				for (int j = 1; j < grvData.Columns.Count; j++)
				{
					
					grvData.Rows[i].Cells[j].ReadOnly = true;

					string tagValue = TextUtils.ToString(grvData.Columns[j].Tag);
					//nếu là những cột tự sinh, là mục kiểm tra
					if (!string.IsNullOrEmpty(tagValue))
					{
						if (qty <= 6)
						{

							if (j <= colResult1.Index)
							{
								grvData.Rows[i].Cells[j].Style.BackColor = _mauHongNhat;

							}

						}
						else
						{
							string tansuat = TextUtils.ToString(grvData.Rows[3].Cells[j].Value);
							if (tansuat.Contains("/"))
							{
								string[] arr = tansuat.Split('/');
								int moc = TextUtils.ToInt(arr[1]);
								// nếu vị trí hàng chia hết cho tần suất. thực hiện đổi màu nền tại hàng đó là LightBlue
								if ((i - 4) % moc == 0)
								{
									grvData.Rows[i].Cells[j].Style.BackColor = _mauHongNhat;

								}
								else
								{
									grvData.Rows[i].Cells[j].Value = "N";
									grvData.Rows[i].Cells[j].Style.BackColor = _mauXam;
									grvData.Rows[i].Cells[j].Style.ForeColor = _mauXam;
									grvData.Rows[i].Cells[j].Selected = false;
									countXam++;
								}
							}
						}

					}
					else
					{
						if (j == colHyp.Index || j == colGear.Index || j == colResult1.Index)
						{
							grvData.Rows[i].Cells[j].Style.BackColor = _mauHongNhat;

						}
					}
				}
				if (countXam == _countGearWorking)
				{
					grvData.Rows[i].Cells[1].Style.BackColor = _mauXam;
					grvData.Rows[i].Cells[1].Selected = false;

					grvData.Rows[i].Cells[2].Style.BackColor = _mauXam;
					grvData.Rows[i].Cells[2].Selected = false;

					grvData.Rows[i].Cells[colResult1.Index].Style.BackColor = _mauXam;
					grvData.Rows[i].Cells[colResult1.Index].Selected = false;

					for (int h = 1; h <= _qtyImage; h++)
					{
						grvData.Rows[i].Cells["colImage" + h].Style.BackColor = _mauXam;
					}

					for (int h = 1; h <= _qtyAudio; h++)
					{
						grvData.Rows[i].Cells["colAudio" + h].Style.BackColor = _mauXam;
					}


				}
				else
				{

				}
			}

			grvData.Focus();
			setFocusByLoadOrder();
		}

		private void txtGear_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				txtQty.Focus();
			}
		}

		private void txtWorkerName_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				txtOrderNo.Focus();
			}
		}


		
        //F8 Lap lai

       /* private void lapLaiConfirm()
        {
			try
			{
				DataGridViewRow currentRow = grvData.CurrentRow;
				int rowIndex = grvData.CurrentRow.Index;
				_SelectedRow = currentRow;
				if (!_isLapLai)
				{
					_isLapLai = true;
					DialogResult rs = MessageBox.Show("Ban co muon lap lai khong ?", "Xac nhan", MessageBoxButtons.YesNo);
					if (rs == DialogResult.No)
					{
						_isLapLai = false;
						SendKeys.Send("{Up}");
						return;
					}
					saveLapLai(rowIndex);
					removeValueLap(rowIndex);
					grvData.CurrentCell = grvData.Rows[rowIndex].Cells[colGear.Index];
					_isLapLai = false;
					SendKeys.Send("{Up}");
				}
			}
			catch { }
           
        }*/


		private void txtQty_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.Handled = !char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar);
		}

		private string getPathFile(DataGridViewCell e)
		{
			try
			{
				if (e.RowIndex < 4) return "";
				if (e.ColumnIndex <= colNumber.Index && e.ColumnIndex >= colQtyLap.Index) return "";
				string order = txtOrderNo.Text.Trim();
				string filename = "";
				if (order == "")
				{
					MessageBox.Show("Bạn chưa nhập order");
					return "";
				}
				string stt = TextUtils.ToString(grvData.Rows[e.RowIndex].Cells[colNumber.Index].Value);
				string qtyLap = TextUtils.ToString(grvData.Rows[e.RowIndex].Cells[colQtyLap.Index].Value);
				DataGridViewColumn col = grvData.Columns[e.ColumnIndex];
				string vtri = TextUtils.ToString(col.Tag);
				filename = order + "_" + stt + "_" + qtyLap + "_" + vtri;
				/*if (e.ColumnIndex > colResult1.Index + _qtyAudio && e.ColumnIndex < colQtyLap.Index)
				{
					filename = order + "_" + stt + "_" + qtyLap + "_" + vtri + ".jpg";
					return filename;
				}

				if (e.OwningColumn.HeaderText.StartsWith("Độ rung động"))
					filename = order + "_" + stt + "_" + qtyLap;*/
				return filename;
			}
			catch
			{
				return "";
			}
		}


		private void cboGear_TextChanged(object sender, EventArgs e)
		{
			ComboBox cbo = (ComboBox)sender;
			if (string.IsNullOrWhiteSpace(cbo.Text.Trim()))
			{
				cbo.BackColor = _mauHongNhat;
			}
			else
			{
				cbo.BackColor = Color.White;
			}
		}


		private void grvData_KeyDown(object sender, KeyEventArgs e)
		{
			try
			{
				if (!_isCatchKey)
				{
					_isCatchKey = true;
					return;
				}

				// get filename in cell
				string filename = "";
				string nameworking = grvData.CurrentCell.OwningColumn.HeaderText;

				int col = grvData.CurrentCell.ColumnIndex;
				int row = grvData.CurrentCell.RowIndex;
				if (col <= colNumber.Index || col >= colQtyLap.Index || row < 4)
				{
					return;
				}
				if (e.KeyCode == Keys.Left)
				{
					filename = getPathFile(grvData[col - 1, row]);
				}
				if (e.KeyCode == Keys.Right)
				{
					filename = getPathFile(grvData[col + 1, row]);
				}
				if (e.KeyCode == Keys.Up)
				{
					filename = getPathFile(grvData[col, row + 1]);
				}
				if (e.KeyCode == Keys.Down)
				{
					filename = getPathFile(grvData[col, row - 1]);
				}

				if (filename == "") return;
			}
			catch { }
			
		}

		private void txtBatch_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) txtHYP.Focus();

		}

		private void txtConfirmer_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) txtBatch.Focus();
		}

		private void txtConfirmer_TextChanged(object sender, EventArgs e)
		{
			TextBox txt = (TextBox)sender;
			if (string.IsNullOrWhiteSpace(txt.Text.Trim()))
			{
				txt.BackColor = _mauHongNhat;
			}
			else
			{
				txt.BackColor = Color.White;
			}
		}

		private void txtBatch_TextChanged(object sender, EventArgs e)
		{
			TextBox txt = (TextBox)sender;
			if (string.IsNullOrWhiteSpace(txt.Text.Trim()))
			{
				txt.BackColor = _mauHongNhat;
			}
			else
			{
				txt.BackColor = Color.White;
			}
		}


		private void tableLayoutPanel1_Paint(object sender, PaintEventArgs e)
		{

		}


        private void panel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void grvData_CellValidated(object sender, DataGridViewCellEventArgs e)
        {
            
        }
        //DataGridViewCellStyle st = new DataGridViewCellStyle()
        private void grvData_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
           
        }

		private void cboLanguage_SelectedIndexChanged(object sender, EventArgs e)
		{
			if (cboLanguage.SelectedIndex == -1) return;

			if (cboLanguage.SelectedIndex == 0) _language = "VN";
			else _language = "JP";
			changeLanguage();
		}
	}
	public class GeneralObject
	{
		public string Confirmer { get; set; }
		public string Batch { get; set; }
		public string OrderCode { get; set; }
		public string HypCode { get; set; }
		public string GearCode { get; set; }
		public string WorkerName { get; set; }
		public int SttStart { get; set; }
		public string Qty { get; set; }
	}
}

