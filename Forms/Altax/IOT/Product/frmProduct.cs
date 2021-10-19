using BMS.Business;
using BMS.Model;
using DevExpress.Utils;
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
	public partial class frmProduct : _Forms
	{
		int _curentNode = 0;
		int _catID = 0;

		public int ProductGroupID { get; private set; }

		public frmProduct()
		{
			InitializeComponent();
		}

		private void frmProduct_Load(object sender, EventArgs e)
		{
			loadTree();
		}
		void loadTree()
		{
			try
			{
				DataTable tbl = TextUtils.Select("Select * from ProductGroup with(nolock) order by ProductGroupCode");

				treeData.DataSource = tbl;
				treeData.KeyFieldName = "ID";
				treeData.PreviewFieldName = "ProductGroupName";
				treeData.ExpandAll();

				DevExpress.XtraTreeList.Nodes.TreeListNode currentNode = treeData.FindNodeByFieldValue("ID", _curentNode);
				treeData.SetFocusedNode(currentNode);
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.ToString());
			}
		}
		private void btnNewGroup_Click(object sender, EventArgs e)
		{
			frmProductGroupDetail frm = new frmProductGroupDetail();
			if (frm.ShowDialog() == DialogResult.OK)
			{
				_curentNode = frm.CurentNode;
				loadTree();
			}
		}

		private void btnEditGroup_Click(object sender, EventArgs e)
		{
			try
			{
				int id = TextUtils.ToInt(treeData.FocusedNode.GetValue(colIDTree));
				if (id == 0) return;
				ProductGroupModel model = (ProductGroupModel)ProductGroupBO.Instance.FindByPK(id);
				frmProductGroupDetail frm = new frmProductGroupDetail();
				frm.ProductGroup = model;
				if (frm.ShowDialog() == DialogResult.OK)
				{
					_curentNode = frm.CurentNode;
					loadTree();
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.ToString());
			}
		}

		private void btnDeleteGroup_Click(object sender, EventArgs e)
		{
			try
			{
				if (treeData.DataSource == null) return;
				int id = (int)treeData.Selection[0].GetValue(colIDTree);
				if (id == 0) return;

				if (ProductBO.Instance.CheckExist("ProductGroupID", id))
				{
					MessageBox.Show("Nhóm sản phẩm này đang được sử dụng nên không thể xóa được.", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}

				if (MessageBox.Show("Bạn có chắc muốn xóa nhóm [" + treeData.Selection[0].GetValue(colNameTree).ToString() + "] không?",
					  TextUtils.Caption, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
					return;
				ProductGroupBO.Instance.Delete(id);
				loadTree();
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message, TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void LoadInfoSearch(int type = 0)
		{
			string[] paraName = new string[2];
			object[] paraValue = new object[2];

			if (type == 0)
			{
				_catID = TextUtils.ToInt(treeData.FocusedNode.GetValue(colIDTree));
			}
			else
			{
				_catID = 0;
			}

			paraName[0] = "@GroupID";
			paraValue[0] = _catID;

			paraName[1] = "@TextFilter";
			paraValue[1] = txtName.Text.Trim();

			ArrayList arr = ProductBO.Instance.GetListObject("spGetProductData", paraName, paraValue);
			grdData.DataSource = arr;
		}

		private void btnAdd_Click(object sender, EventArgs e)
		{
			int id = TextUtils.ToInt(treeData.FocusedNode.GetValue(colIDTree));
			frmProductDetail frm = new frmProductDetail();
			frm.ProductGroupID = id;
			if (frm.ShowDialog() == DialogResult.OK)
			{
				LoadInfoSearch();
			}
		}
		private void btnCopy_Click(object sender, EventArgs e)
		{

		}

		int _rownIndex = 0;
		private void btnEdit_Click(object sender, EventArgs e)
		{
			int id = TextUtils.ToInt(grvData.GetFocusedRowCellValue(colID));
			if (id == 0) return;
			ProductModel model = (ProductModel)ProductBO.Instance.FindByPK(id);
			int catId = TextUtils.ToInt(grvData.GetFocusedRowCellValue(colGroup));
			_rownIndex = grvData.FocusedRowHandle;

			frmProductDetail frm = new frmProductDetail();
			frm.ProductGroupID = catId;
			frm.productModel = model;
			if (frm.ShowDialog() == DialogResult.OK)
			{
				LoadInfoSearch();
			}
		}

		private void btnDel_Click(object sender, EventArgs e)
		{
			if (!grvData.IsDataRow(grvData.FocusedRowHandle))
				return;

			int strID = TextUtils.ToInt(grvData.GetFocusedRowCellValue("ID"));

			string strName = TextUtils.ToString(grvData.GetFocusedRowCellValue("ProductCode"));

			//if (ProductStepBO.Instance.CheckExist("ProductID", strID))
			//{
			//    MessageBox.Show("Sản phẩm này đang được sử dụng ở nghiệp vụ khác nên không thể xóa được!", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
			//    return;
			//}

			//if (ProductWorkingBO.Instance.CheckExist("ProductID", strID))
			//{
			//    MessageBox.Show("Sản phẩm này đang được sử dụng ở nghiệp vụ khác nên không thể xóa được!", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
			//    return;
			//}

			if (ProductCheckHistoryDetailBO.Instance.CheckExist("ProductID", strID))
			{
				MessageBox.Show("Sản phẩm này đã có lịch sử kiểm tra nên không thể xóa được!", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if (MessageBox.Show(String.Format("Bạn có chắc muốn xóa sản phẩm [{0}] không?", strName), TextUtils.Caption, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
			{
				try
				{
					ProductStepBO.Instance.DeleteByAttribute("ProductID", strID);
					ProductWorkingBO.Instance.DeleteByAttribute("ProductID", strID);
					//ProductCheckHistoryDetailBO.Instance.DeleteByAttribute("",);
					ProductBO.Instance.Delete(strID);
					grvData.DeleteRow(grvData.FocusedRowHandle);
				}
				catch (Exception ex)
				{
					MessageBox.Show("Có lỗi xảy ra khi thực hiện thao tác, xin vui lòng thử lại sau.\n" + ex.ToString());
				}
			}
		}

		private void btnExcel_Click(object sender, EventArgs e)
		{
			try
			{
				if (grvData.RowCount > 0)
				{
					FolderBrowserDialog od = new FolderBrowserDialog();
					if (od.ShowDialog() == DialogResult.OK)
					{
						TextUtils.ExportExcel(grvData, od.SelectedPath, string.Format("DanhSachSanPham_{0}", DateTime.Now.ToString("yyyy_MM_dd_HH_mm_ss")));
					}
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.ToString());
			}
		}

		private void btnFindAll_Click(object sender, EventArgs e)
		{
			LoadInfoSearch(1);
		}

		private void treeData_FocusedNodeChanged(object sender, DevExpress.XtraTreeList.FocusedNodeChangedEventArgs e)
		{
			try
			{
				int id = TextUtils.ToInt(treeData.FocusedNode.GetValue(colIDTree));
				if (id > 0)
				{
					txtName.Text = txtName.Text = "";
					LoadInfoSearch();
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.ToString());
			}
		}

		private void txtName_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				LoadInfoSearch(1);
			}
		}

		private void grvData_DoubleClick(object sender, EventArgs e)
		{
			if (grvData.RowCount > 0 && btnEdit.Enabled == true)
				btnEdit_Click(null, null);
		}

		private void btnShowWorking_Click(object sender, EventArgs e)
		{
			frmProductWorking frm = new frmProductWorking();
			frm.ProductID = TextUtils.ToInt(grvData.GetFocusedRowCellValue(colID));
			frm.ProductCode = TextUtils.ToString(grvData.GetFocusedRowCellValue(colCode));
			frm.Show();
		}

		private void btnImportExcel_Click(object sender, EventArgs e)
		{
			frmImportProductExcel frm = new frmImportProductExcel();
			frm.Show();
		}

		private void btnSungLuc_Click(object sender, EventArgs e)
		{
			frmProductForceGun frm = new frmProductForceGun();
			frm.Show();
		}

		private void btnUpdateStepCode_Click(object sender, EventArgs e)
		{
			using (WaitDialogForm fWait = new WaitDialogForm("", ""))
			{
				TextUtils.ExcuteSQL("exec spReUpdateProductWorkingStepCode");
			}
		}

		private void btnSpeedReducer_Click(object sender, EventArgs e)
		{
			frmReduceSpeed frm = new frmReduceSpeed();
			frm.Show();
		}
	}
}
