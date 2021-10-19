using DevExpress.Utils;
using HP.Business;
using HP.Model;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace BMS
{
	public partial class frmKnifeGroupDetailH : _Forms
	{
		public KnifeGroupModel KnifeGroup = new KnifeGroupModel();
		public int Loai;
		ArrayList arrStep = new ArrayList();

		public int CurentNode { get; set; }
		string _Name = "";
		int _STT = 0;
		string _Des = "";

		public frmKnifeGroupDetailH()
		{
			InitializeComponent();
		}

		private void frmProductGroup_Load(object sender, EventArgs e)
		{
			txtCode.Text = KnifeGroup.KnifeGroupCode;
			txtName.Text = KnifeGroup.KnifeGroupName;

			loadData();

			DataTable dt = TextUtilsHP.Select("SELECT * FROM [HypoidPinion].[dbo].[KnifeGroup]");
			cbCopy.Properties.DataSource = dt;
			cbCopy.Properties.DisplayMember = "KnifeGroupName";
			cbCopy.Properties.ValueMember = "ID";
		}

		#region Methods
		private void loadData()
		{
			try
			{
				arrStep = WorkingStepBO.Instance.GetListObject("spGetWorkingStep_ByGroupID", new string[] { "@GroupID" }, new object[] { KnifeGroup.ID });
				grdData.DataSource = arrStep;
			}
			catch (Exception)
			{
			}
		}

		#endregion

		#region Buttons Events
		private void btnNew_Click(object sender, EventArgs e)
		{
			frmKnifeWorkingStepDetailH frm = new frmKnifeWorkingStepDetailH();
			frm.ArrStep = arrStep;
			if (frm.ShowDialog() == DialogResult.OK)
			{
				//arrStep.Insert(0, frm.WorkingStep);
				arrStep.Add(frm.WorkingStep);
				grdData.DataSource = null;
				grdData.DataSource = arrStep;
			}
		}

		private void btnEdit_Click(object sender, EventArgs e)
		{
			try
			{
				int id = TextUtilsHP.ToInt(grvData.GetFocusedRowCellValue(colID));
				int sortOrder = TextUtilsHP.ToInt(grvData.GetFocusedRowCellValue(colSortOrder));
				WorkingStepModel model = new WorkingStepModel();
				var dr = grvData.GetRow(grvData.FocusedRowHandle);
				model = (WorkingStepModel)dr;
				frmKnifeWorkingStepDetailH frm = new frmKnifeWorkingStepDetailH();
				frm.WorkingStep = model;
				frm.ArrStep = arrStep;
				if (frm.ShowDialog() == DialogResult.OK)
				{
					var lst = arrStep.ToArray().ToList();
					lst = lst.Where(o => ((WorkingStepModel)o).ID != id && ((WorkingStepModel)o).SortOrder != sortOrder).ToList();
					arrStep = new ArrayList();
					for (int i = 0; i < lst.Count; i++)
					{
						arrStep.Add(lst[i]);
					}

					arrStep.Add(frm.WorkingStep);

					grdData.DataSource = null;
					grdData.DataSource = arrStep;
				}
			}
			catch (Exception ex)
			{
			}
		}

		private void btnDelete_Click(object sender, EventArgs e)
		{
			if (grvData.FocusedRowHandle < 0)
				return;

			int strID = TextUtilsHP.ToInt(grvData.GetFocusedRowCellValue("ID"));

			string strName = TextUtilsHP.ToString(grvData.GetFocusedRowCellValue("SortOrder"));

			if (WorkingBO.Instance.CheckExist("WorkingStepID", strID))
			{
				MessageBox.Show("Công đoạn này đang được sử dụng bên Mục kiểm tra nên không thể xóa được!", TextUtilsHP.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if (MessageBox.Show(String.Format("Bạn có chắc muốn xóa công đoạn [{0}] không?", strName), TextUtilsHP.Caption, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
			{
				try
				{
					if (strID > 0)
					{
						TextUtilsHP.ExcuteProcedure("spUpdateKnifeStep_ByWorkingStep"
						, new string[] { "@WorkingStepID", "@Type" }
						, new object[] { strID, 3 });
					}

					grvData.DeleteSelectedRows();
				}
				catch
				{
					MessageBox.Show("Có lỗi xảy ra khi thực hiện thao tác, xin vui lòng thử lại sau.");
				}
			}
		}

		private void btnSave_Click(object sender, EventArgs e)
		{
			try
			{
				using (WaitDialogForm fWait = new WaitDialogForm("Please Wait...", "Processing Data"))
				{
					KnifeGroup.KnifeGroupCode = txtCode.Text.Trim();
					KnifeGroup.KnifeGroupName = txtName.Text.Trim();
					KnifeGroup.Type = Loai;
					if (KnifeGroup.ID <= 0)
					{
						KnifeGroup.ID = (int)KnifeGroupBO.Instance.Insert(KnifeGroup);
					}
					else
					{
						KnifeGroupBO.Instance.Update(KnifeGroup);
					}
					CurentNode = KnifeGroup.ID;
					for (int i = 0; i < arrStep.Count; i++)
					{
						_STT = TextUtilsHP.ToInt(grvData.GetRowCellValue(i, colSortOrder));
						_Name = TextUtilsHP.ToString(grvData.GetRowCellValue(i, colName));
						_Des = TextUtilsHP.ToString(grvData.GetRowCellValue(i, colDes));
						WorkingStepModel step = (WorkingStepModel)arrStep[i];
						int type = 0;
						if (step.ID <= 0)
						{
							step.KnifeGroupID = KnifeGroup.ID;
							step.ID = (int)WorkingStepBO.Instance.Insert(step);
							type = 1;
							//Thêm step vào các sản phẩm thuộc nhóm này
						}
						else
						{
							//string sql= $"SELECT TOP 1 ws.ID FROM dbo.WorkingStep ws JOIN dbo.ProductStep ps ON ws.ID = ps.WorkingStepID WHERE WorkingStepID = {step.ID} and ws.WorkingStepCode = N'{_Name}'  and ws.Description = N'{_Des}' and ws.SortOrder ='{_STT}'";
							//int id = TextUtilsHP.ToInt(TextUtilsHPIE.ExcuteScalar(sql));
							//if (id == 0)
							//{
							WorkingStepBO.Instance.Update(step);
							type = 2;
							// type = 1;
							//Cập nhật vào các sản phẩm thuộc nhóm này
							//}
						}

						TextUtilsHP.ExcuteProcedure("spUpdateProductStep_ByWorkingStep"
							  , new string[] { "@WorkingStepID", "@Type" }
							  , new object[] { step.ID, type });

					}

					DialogResult = DialogResult.OK;
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message, TextUtilsHP.Caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
			}


		}

		private void btnClose_Click(object sender, EventArgs e)
		{
			this.Close();
		}
		#endregion

		private void grdData_DoubleClick(object sender, EventArgs e)
		{
			if (grvData.RowCount > 0 && btnEdit.Enabled == true)
				btnEdit_Click(null, null);
		}

		private void button1_Click(object sender, EventArgs e)
		{
			WorkingStepModel model = new WorkingStepModel();
			for (int i = 0; i < grvData.RowCount; i++)
			{
				model.WorkingStepCode = TextUtilsHP.ToString(grvData.GetRowCellValue(i, colName));
				model.SortOrder = TextUtilsHP.ToInt(grvData.GetRowCellValue(i, colSortOrder));
				model.Description = TextUtilsHP.ToString(grvData.GetRowCellValue(i, colDes));
				model.CreatedDate = DateTime.Now;
				model.KnifeGroupID = TextUtilsHP.ToInt(cbCopy.EditValue);
				WorkingStepBO.Instance.Insert(model);
			}

		}
		private void btnClose_Click_1(object sender, EventArgs e)
		{
			this.Close();
		}

		private void frmProductGroupDetailH_FormClosed(object sender, FormClosedEventArgs e)
		{
			DialogResult = DialogResult.OK;
		}
	}
}
