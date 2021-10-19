using HP.Business;
using HP.Model;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace BMS
{
	public partial class frmWorkingDaoJig : _Forms
	{
		public frmWorkingDaoJig()
		{
			InitializeComponent();
		}

		private void frmWorking_Load(object sender, EventArgs e)
		{
			cboLoai.SelectedIndex = 0;
		}

		private void btnNew_Click(object sender, EventArgs e)
		{
			frmWorkingDaoJigDetail frm = new frmWorkingDaoJigDetail();
			frm._Loai = cboLoai.SelectedIndex;
			frm._GroupID = TextUtils.ToInt(cboGroup.SelectedValue);
			frm._StepID = TextUtils.ToInt(treeDataCD.FocusedNode.GetValue(colIDTree));
			if (frm.ShowDialog() == DialogResult.OK)
			{
				LoadTreeData();
			}
		}

		private void btnEdit_Click(object sender, EventArgs e)
		{
			try
			{
				int id = TextUtils.ToInt(grvDetail.GetFocusedRowCellValue(colWorkingID));
				if (id == 0) return;
				WorkingModel model = (WorkingModel)WorkingBO.Instance.GetObjectByID(id);
				frmWorkingDaoJigDetail frm = new frmWorkingDaoJigDetail();
				frm._Loai = cboLoai.SelectedIndex;
				frm._WorkingModel = model;
				if (frm.ShowDialog() == DialogResult.OK)
				{
					LoadTreeData();
				}
			}
			catch (Exception)
			{
			}
		}
		private void grdDetail_DoubleClick(object sender, EventArgs e)
		{
			if (grvDetail.RowCount > 0 && btnEdit.Enabled == true)
				btnEdit_Click(null, null);
		}

		private void btnDelete_Click(object sender, EventArgs e)
		{
			try
			{
				int ParentID = TextUtils.ToInt(treeDataCD.FocusedNode.GetValue(colIDTree));
				int SortOrder = TextUtils.ToInt(grvDetail.GetFocusedRowCellValue(colSortOrder));
				string CD = TextUtils.ToString(grvDetail.GetFocusedRowCellValue(colProductStepCode));
				int id = TextUtils.ToInt(grvDetail.GetFocusedRowCellValue(colWorkingID));
				if (id == 0) return;

				if (MessageBox.Show("Bạn có chắc muốn xóa mục cần kiểm tra [" + grvDetail.GetFocusedRowCellValue(colWorkingName).ToString() + "] không?",
					  TextUtils.Caption, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
					return;
				WorkingBO.Instance.Delete(id);
				grvDetail.DeleteSelectedRows();
				//Xóa các thằng con thuộc thằng cha này đi
				TextUtilsHP.ExcuteSQL(string.Format("EXEC spUpdateMultiProductWorking {0}, 2", id));
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message, TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void cboLoai_SelectedIndexChanged(object sender, EventArgs e)
		{
			//Load Group
			LoadGroup();
			if (cboLoai.SelectedIndex == 1)
			{
				colFrequency.Visible = true;
			}
			else
			{
				colFrequency.Visible = false;
			}	
		}
		void LoadGroup()
		{
			DataTable dt = TextUtilsHP.Select($"SELECT ID, KnifeGroupCode FROM [HypoidPinion].[dbo].[KnifeGroup] WHERE Type =N'{cboLoai.SelectedIndex}'");
			DataRow dr = dt.NewRow();
			dr["ID"] = 0;
			dr["KnifeGroupCode"] = "";
			dt.Rows.InsertAt(dr, 0);
			cboGroup.DisplayMember = "KnifeGroupCode";
			cboGroup.ValueMember = "ID";
			cboGroup.DataSource = dt;
		}
		void LoadTreeData()
		{
			if (cboGroup.SelectedIndex == 0) return;

			DataTable tbl = TextUtilsHP.LoadDataFromSP("spLoadTreeDataCD", "A", new string[] { "@KnifeGroupID" }, new object[] { TextUtilsHP.ToInt(cboGroup.SelectedValue) });
			treeDataCD.DataSource = tbl;
			treeDataCD.KeyFieldName = "ID";
			treeDataCD.PreviewFieldName = "KnifeCode";
			treeDataCD.ExpandAll();

			//DevExpress.XtraTreeList.Nodes.TreeListNode currentNode = treeData.FindNodeByFieldValue("ID", _curentNode);
			//treeData.SetFocusedNode(currentNode);
		}

		private void btnCopyGroupProducts_Click(object sender, EventArgs e)
		{

		}
		private void treeData_FocusedNodeChanged(object sender, DevExpress.XtraTreeList.FocusedNodeChangedEventArgs e)
		{
			int id = TextUtils.ToInt(treeDataCD.FocusedNode.GetValue(colIDTree));
			if (id > 0)
			{
				LoadInfoSearch(id);
			}
		}
		void LoadInfoSearch(int id)
		{
			DataTable dt = TextUtilsHP.LoadDataFromSP("spLoadWorkingName", "LoadWorkingName", new string[] { "@WorkingStepID", "@Chk" }, new object[] { id, chkAll.Checked });
			grdDetail.DataSource = dt;
		}
		private void chkAll_CheckedChanged(object sender, EventArgs e)
		{
			if (treeDataCD.Nodes.Count <= 0) return;
			int id = TextUtils.ToInt(treeDataCD.FocusedNode.GetValue(colIDTree));
			if (id > 0)
			{
				LoadInfoSearch(id);
			}
		}

		private void cboGroup_SelectedIndexChanged(object sender, EventArgs e)
		{
			//Load công đoạn theo nhóm
			LoadTreeData();
		}
	}
}
