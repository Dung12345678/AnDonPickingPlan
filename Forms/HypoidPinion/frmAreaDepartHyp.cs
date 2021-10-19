using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using HP.Business;
using HP.Model;

namespace BMS
{
	public partial class frmAreaDepartHyp : _Forms
	{
		#region Variables
		int prevRow;
		int prevRow1;
		bool showAll = false;
		#endregion

		#region Method
		public frmAreaDepartHyp()
		{
			InitializeComponent();
		}

		void LoadArea()
		{
			DataTable arr = TextUtils.Select(cGlobalFunction.GetAllQuery("Area", "ID"));
			dtgvArea.DataSource = arr;
		}
		void LoadDepartment()
		{
			if (cbShowAll.Checked)
				showAll = true;
			else
				showAll = false;
			int id = TextUtils.ToInt(gvArea.GetFocusedRowCellValue(colIDAr));
			if (id == 0) return;
			DataTable arr = TextUtils.LoadDataFromSP("spGetDepartmentListWithFilter", "a", new string[] { "@keyword" , "@areaID", "@showAll" }, new object[] { txbSearch.Text.Trim(), id, showAll });
			dtgvDepartment.DataSource = arr;
		}
		#endregion
		
		#region Event
		private void frmAreaDepartHyp_Load(object sender, EventArgs e)
		{
			LoadArea();
			LoadDepartment();
		}
		private void btnEditArea_Click(object sender, EventArgs e)
		{
			int id = TextUtils.ToInt(gvArea.GetFocusedRowCellValue(colIDAr));
			if (id == 0) return;
			prevRow = gvArea.GetSelectedRows()[0];
			AreaModel model = (AreaModel)AreaBO.Instance.FindByPK(id);

			frmAddEditAreas frm = new frmAddEditAreas(cGlobalVariables.Edit);
			frm.areaModel = model;
			if (frm.ShowDialog() == DialogResult.OK)
			{
				LoadArea();
				LoadDepartment();
				gvArea.FocusedRowHandle = prevRow;
			}
		}

		private void btnCreateArea_Click(object sender, EventArgs e)
		{
			frmAddEditAreas frm = new frmAddEditAreas(cGlobalVariables.Add);
			if (frm.ShowDialog() == DialogResult.OK)
			{
				LoadArea();
			}
		}

		private void btnDelArea_Click(object sender, EventArgs e)
		{
			if (!gvArea.IsDataRow(gvArea.FocusedRowHandle))
				return;
			int strID = TextUtils.ToInt(gvArea.GetFocusedRowCellValue(colIDAr));
			string str = TextUtils.ToString(gvArea.GetFocusedRowCellValue("AreaCode"));

			try
			{
				if (MessageBox.Show(String.Format("Bạn có chắc muốn xóa khu vực [{0}] không?", str), TextUtils.Caption, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
				{
					if (DepartmentBO.Instance.CheckExist("AreaID", strID))
					{
						MessageBox.Show("Khu vực này đã tồn tại các bộ phận trực thuộc!", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
						return;
					}
					else
					{
						AreaBO.Instance.Delete(strID);
						gvArea.DeleteRow(gvArea.FocusedRowHandle);
						MessageBox.Show("Xoá thành công", TextUtils.Caption);
					}
				}

			}
			catch (Exception ex)
			{
				MessageBox.Show("Lỗi hệ thống!", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
		}

		private void btnRefresh_Click(object sender, EventArgs e)
		{
			LoadArea();
			LoadDepartment();
		}

		private void btnCreateDep_Click(object sender, EventArgs e)
		{
			frmAddEditDepartment frm = new frmAddEditDepartment(cGlobalVariables.Add);
			if (frm.ShowDialog() == DialogResult.OK) 
			{
				LoadDepartment();
			}
		}

		private void btnEditDep_Click(object sender, EventArgs e)
		{
			int id = TextUtils.ToInt(gvDepartment.GetFocusedRowCellValue(colIDDep));
			if (id == 0) return;
			prevRow1 = gvDepartment.GetSelectedRows()[0];
			DepartmentModel model = (DepartmentModel)DepartmentBO.Instance.FindByPK(id);

			frmAddEditDepartment frm = new frmAddEditDepartment(cGlobalVariables.Edit);
			frm.departmentModel = model;
			if (frm.ShowDialog() == DialogResult.OK)
			{
				LoadDepartment();
				gvDepartment.FocusedRowHandle = prevRow1;
			}
		}

		private void btnDelDep_Click(object sender, EventArgs e)
		{
			if (!gvDepartment.IsDataRow(gvDepartment.FocusedRowHandle))
				return;
			int strID = TextUtils.ToInt(gvDepartment.GetFocusedRowCellValue(colIDDep));
			string str = TextUtils.ToString(gvDepartment.GetFocusedRowCellValue("DepartmentCode"));

			try
			{
				if (MessageBox.Show(String.Format("Bạn có chắc muốn xóa bộ phận [{0}] không?", str), TextUtils.Caption, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
				{
					if (StageBO.Instance.CheckExist("DepartmentID", strID))
					{
						MessageBox.Show("Bộ phận này đã tồn tại các công đoạn trực thuộc!", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
						return;
					}
					else
					{
						DepartmentBO.Instance.Delete(strID);
						gvDepartment.DeleteRow(gvDepartment.FocusedRowHandle);
						MessageBox.Show("Xoá thành công", TextUtils.Caption);
					}
				}

			}
			catch (Exception ex)
			{
				MessageBox.Show("Lỗi hệ thống!", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
		}


		private void btnSearch_Click(object sender, EventArgs e)
		{
			LoadDepartment();
		}

		private void btnCancelSearch_Click(object sender, EventArgs e)
		{
			txbSearch.Text = "";
			LoadDepartment();
		}

		private void cbShowAll_CheckedChanged(object sender, EventArgs e)
		{
			LoadDepartment();
		}

		private void gvArea_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
		{
			LoadDepartment();
		}

		private void txbSearch_KeyPress(object sender, KeyPressEventArgs e)
		{
			if (e.KeyChar == (char)13)
				btnSearch_Click(null, null);
		}

		private void gvDepartment_DoubleClick(object sender, EventArgs e)
		{
			btnEditDep_Click(null, null);
		}

		private void dtgvArea_DoubleClick(object sender, EventArgs e)
		{
			btnEditArea_Click(null, null);
		}

		private void btnCheckStage_Click(object sender, EventArgs e)
		{

		}
		#endregion

	}
}
