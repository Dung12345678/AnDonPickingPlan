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
	public partial class frmKnifeList : _Forms
	{
		#region Variables
		int prevRow;
		int _curentNode = 0;
		int _catID = 0;
		#endregion

		#region Methods
		public frmKnifeList()
		{
			InitializeComponent();
		}
		void LoadData()
		{
			DataTable arr = TextUtilsHP.GetDataTableFromSP("spLoadKnife", new string[] { "@Text", "@Chk", "@GroupKnifeID" }, new object[] { txtFindDate.Text.Trim(), 1, TextUtils.ToInt(treeDataGroup.FocusedNode.GetValue(colIDTreeGroup)) });
			treeDataKnife.DataSource = arr;
			treeDataKnife.KeyFieldName = "ID";
			treeDataKnife.PreviewFieldName = "KnifeCode";
			treeDataKnife.ExpandAll();
		}

		void LoadUnavailableData()
		{
			DataTable arr = TextUtilsHP.GetDataTableFromSP("spLoadKnife", new string[] { "@Text", "@Chk", "@GroupKnifeID" }, new object[] { txtFindDate.Text.Trim(), 0, TextUtils.ToInt(treeDataGroup.FocusedNode.GetValue(colIDTreeGroup)) });
			treeDataKnife.DataSource = arr;
			treeDataKnife.KeyFieldName = "ID";
			treeDataKnife.PreviewFieldName = "KnifeCode";
			treeDataKnife.ExpandAll();
		}
		#endregion

		#region Events
		private void btnCreateKnife_Click(object sender, EventArgs e)
		{
			//Dao
			int id = TextUtils.ToInt(treeDataGroup.FocusedNode.GetValue(colIDTreeGroup));
			if (cboLoai.SelectedIndex == 0)
			{

				frmAddEditKnife frm = new frmAddEditKnife(cGlobalVariables.Add);
				frm.KnifeGroupID = id;
				frm.Loai = cboLoai.SelectedIndex;
				if (frm.ShowDialog() == DialogResult.OK)
				{
					LoadInfoSearch();
				}
			}
			//Jig
			else if (cboLoai.SelectedIndex == 1)
			{
				frmAddEditJig frm = new frmAddEditJig();
				frm.KnifeGroupID = id;
				frm.Loai = cboLoai.SelectedIndex;
				if (frm.ShowDialog() == DialogResult.OK)
				{
					LoadInfoSearch();
				}
			}
			//khuôn
			else if (cboLoai.SelectedIndex == 2)
			{
				frmShapeDetails frm = new frmShapeDetails();
				frm.KnifeGroupID = id;
				frm.Loai = cboLoai.SelectedIndex;
				if (frm.ShowDialog() == DialogResult.OK)
				{
					//LoadAll();
					LoadInfoSearch();
				}
			}
			else
			{
				frmAddEditTools frm = new frmAddEditTools();
				frm.KnifeGroupID = id;
				frm.Loai = cboLoai.SelectedIndex;
				if (frm.ShowDialog() == DialogResult.OK)
				{
					//Load lại dụng cụ đo
					LoadToods();
					//LoadInfoSearch();
				}
			}
		}

		private void btnEditKnife_Click(object sender, EventArgs e)
		{
			if (treeDataKnife.DataSource == null) return;
			if (cboLoai.SelectedIndex == 0) //Sửa Dao
			{
				int id = TextUtilsHP.ToInt(treeDataKnife.FocusedNode.GetValue(colID));
				if (id == 0) return;
				KnifeDetailListModel model = (KnifeDetailListModel)KnifeDetailListBO.Instance.FindByPK(id);

				frmAddEditKnife frm = new frmAddEditKnife(cGlobalVariables.Edit);
				frm.knifeDetailListModel = model;
				frm.Loai = cboLoai.SelectedIndex;
				if (frm.ShowDialog() == DialogResult.OK)
				{
					//cbShowUnavailable_CheckedChanged(null, null);
					LoadInfoSearch();
				}
			}
			else if (cboLoai.SelectedIndex == 1) //Sửa Jig
			{
				int id = TextUtilsHP.ToInt(treeDataKnife.FocusedNode.GetValue(colID));
				if (id == 0) return;
				KnifeDetailListModel model = (KnifeDetailListModel)KnifeDetailListBO.Instance.FindByPK(id);
				frmAddEditJig frm = new frmAddEditJig();
				frm._KnifeDetailListModel = model;
				frm.Loai = cboLoai.SelectedIndex;
				if (frm.ShowDialog() == DialogResult.OK)
				{
					//LoadAll();
					LoadInfoSearch();
				}
			}
			else if (cboLoai.SelectedIndex == 2) // Sửa khuôn
			{
				int id = TextUtilsHP.ToInt(treeDataKnife.FocusedNode.GetValue(colID));
				if (id == 0) return;
				KnifeDetailListModel model = (KnifeDetailListModel)KnifeDetailListBO.Instance.FindByPK(id);
				frmShapeDetails frm = new frmShapeDetails();
				frm.Loai = cboLoai.SelectedIndex;
				frm._KnifeDetailListModel = model;
				if (frm.ShowDialog() == DialogResult.OK)
				{
					//LoadData();
					LoadInfoSearch();
				}
			}
			else
			{
				//Sửa dụng cụ đo
				int id = TextUtilsHP.ToInt(treeDataKnife.FocusedNode.GetValue(colID));
				if (id == 0) return;
				ToolsModel model = (ToolsModel)ToolsBO.Instance.FindByPK(id);
				frmAddEditTools frm = new frmAddEditTools();
				frm.Loai = cboLoai.SelectedIndex;
				frm._toolsModel = model;
				if (frm.ShowDialog() == DialogResult.OK)
				{
					LoadToods();
				}
			}
		}

		private void btnDelKnife_Click(object sender, EventArgs e)
		{
			if (treeDataKnife.DataSource == null) return;
			if (treeDataKnife.AllNodesCount <= 0) return;
			int strID = TextUtilsHP.ToInt(treeDataKnife.FocusedNode.GetValue(colID));
			string str = TextUtilsHP.ToString(treeDataKnife.FocusedNode.GetValue("KnifeCode"));

			try
			{
				if (MessageBox.Show(String.Format("Bạn có chắc muốn xóa dụng cụ [{0}] không?", str), TextUtilsHP.Caption, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
				{
					if (cboLoai.SelectedIndex == 0)
					{
						if (KnifeSharpeningDetailsBO.Instance.CheckExist("KnifeID", strID))
						{
							MessageBox.Show("Dao đã được sử dụng! Vui lòng chọn hủy dao thay cho xóa", TextUtilsHP.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
							return;
						}
						else
						{
							KnifeDetailListBO.Instance.Delete(strID);
							treeDataKnife.DeleteSelectedNodes();
						}
					}
					else if (cboLoai.SelectedIndex == 3)
					{
						ToolsBO.Instance.Delete(strID);
						treeDataKnife.DeleteSelectedNodes();
					}
					else
					{
						KnifeDetailListBO.Instance.Delete(strID);
						treeDataKnife.DeleteSelectedNodes();
					}
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show("Lỗi hệ thống!", TextUtilsHP.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
		}

		private void btnRefresh_Click(object sender, EventArgs e)
		{
			cbShowUnavailable_CheckedChanged(null, null);
		}

		private void frmKnifeList_Load(object sender, EventArgs e)
		{
			cboLoai.SelectedIndex = 0;

			//LoadData();
		}
		private void gvKnife_DoubleClick(object sender, EventArgs e)
		{
			btnEditKnife_Click(null, null);
		}

		private void cbShowUnavailable_CheckedChanged(object sender, EventArgs e)
		{
			if (cbShowUnavailable.Checked)
			{
				//LoadUnavailableData();
			}
			else
			{
				//LoadData();
			}
		}

		private void btnDisposeTool_Click(object sender, EventArgs e)
		{
			if (treeDataKnife.AllNodesCount <= 0) return;
			TextUtilsHP.ToInt(treeDataKnife.FocusedNode.GetValue(colParentID));
			int id = TextUtilsHP.ToInt(treeDataKnife.FocusedNode.GetValue(colID));
			if (id == 0) return;

			frmKnifeDisposed frm = new frmKnifeDisposed();
			frm.knifeID = id;
			if (frm.ShowDialog() == DialogResult.OK)
			{
				cbShowUnavailable_CheckedChanged(null, null);
			}
		}
		private void btnSharpenKnife_Click(object sender, EventArgs e)
		{
			//if (treeData.AllNodesCount <= 0) return;
			//int id = TextUtilsHP.ToInt(treeData.FocusedNode.GetValue(colID));
			//if (id == 0) return;

			frmKnifeSharpen frm = new frmKnifeSharpen();
			//frm.knifeID = id;
			if (frm.ShowDialog() == DialogResult.OK)
			{
				cbShowUnavailable_CheckedChanged(null, null);
			}
		}

		private void btnProcessTool_Click(object sender, EventArgs e)
		{
			//if (treeData.AllNodesCount <= 0) return;
			//int id = TextUtilsHP.ToInt(treeData.FocusedNode.GetValue(colID));
			string KnifeCode = TextUtilsHP.ToString(treeDataKnife.FocusedNode.GetValue(colKnifeCode));
			//if (id == 0) return;

			frmKnifeProcessedList frm = new frmKnifeProcessedList();
			//frm.knifeID = id;
			frm._KnifeCode = KnifeCode;
			if (frm.ShowDialog() == DialogResult.OK)
			{
				cbShowUnavailable_CheckedChanged(null, null);
			}
		}
		private void btnProcessChart_Click(object sender, EventArgs e)
		{
			frmKnifeProcessedChart frm = new frmKnifeProcessedChart();
			frm.ShowDialog();
		}

		#endregion

		private void btnFindDate_Click(object sender, EventArgs e)
		{
			if (cboLoai.SelectedIndex == 0)
			{
				if (cbShowUnavailable.Checked)
				{
					LoadUnavailableData();
				}
				else
				{
					LoadData();
				}
			}
			else if (cboLoai.SelectedIndex == 3)
			{
				LoadToods();
			}
			else
			{
				LoadInfoSearch(0);
			}

		}

		private void cboLoai_SelectedIndexChanged(object sender, EventArgs e)
		{
			//Load Group theo nhóm
			if (cboLoai.SelectedIndex != 0)
			{
				btnDisposeTool.Visible = false;
				toolStripSeparator3.Visible = false;
			}
			else
			{
				btnDisposeTool.Visible = true;
				toolStripSeparator3.Visible = true;
			}
			if (cboLoai.SelectedIndex == 1 || cboLoai.SelectedIndex == 3)
			{
				colFrequencyCheck.Visible = true;
				colFrequencyCheck.VisibleIndex = 12;
				colDateOld.Visible = true;
				colDateOld.VisibleIndex = 13;
			}
			else
			{
				colFrequencyCheck.Visible = false;
				colDateOld.Visible = false;
			}
			treeDataKnife.DataSource = null;
			loadTreeGroup();

		}
		void loadTreeGroup()
		{
			try
			{
				DataTable tbl = TextUtilsHP.Select($"Select * from KnifeGroup with(nolock) Where Type={cboLoai.SelectedIndex} order by KnifeGroupCode");

				treeDataGroup.DataSource = tbl;
				treeDataGroup.KeyFieldName = "ID";
				treeDataGroup.PreviewFieldName = "KnifeGroupName";
				treeDataGroup.ExpandAll();

				DevExpress.XtraTreeList.Nodes.TreeListNode currentNode = treeDataGroup.FindNodeByFieldValue("ID", _curentNode);
				treeDataGroup.SetFocusedNode(currentNode);
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.ToString());
			}
		}
		void Visible1(bool t)
		{
			colKnifeCode.Visible = t;
			colKnifeName.Visible = t;
			colDescription.Visible = t;
			colFullName.Visible = t;
			colImportedDate.Visible = t;
			if(t!=false)
			{
				colImportedDate.VisibleIndex = 6;
				colFullName.VisibleIndex = 5;
				colDescription.VisibleIndex = 2;
				colKnifeName.VisibleIndex = 1;
				colKnifeCode.VisibleIndex = 0;
			}
		}
		void Visible2(bool t)
		{
			colCode.Visible = t;
			colName.Visible = t;
			colStdTools.Visible = t;
			colMin.Visible = t;
			colMax.Visible = t;
			if (t != false)
			{
				colCode.VisibleIndex = 0;
				colName.VisibleIndex = 1;
				colStdTools.VisibleIndex = 2;
				colMin.VisibleIndex = 3;
				colMax.VisibleIndex = 4;
			}
		}
		void LoadToods()
		{
			if (treeDataGroup.AllNodesCount <= 0) return;
			_catID = TextUtils.ToInt(treeDataGroup.FocusedNode.GetValue(colIDTreeGroup));
			DataTable arr = TextUtilsHP.GetDataTableFromSP(
											"spLoadTools",
											new string[] { "@GroupID" },
											new object[] { _catID });
			treeDataKnife.DataSource = arr;
		}
		//void LoadAll()
		//{
		//	DataTable arr = TextUtilsHP.GetDataTableFromSP(
		//									"spLoadAll",
		//									new string[] { "@DateStart", "@DateEnd", "@Text", "@Type" },
		//									new object[]{dtpFrom.Value.ToString("yyyy/MM/dd HH:mm:ss")
		//												, dtpTo.Value.ToString("yyyy/MM/dd HH:mm:ss")
		//												, txtFindDate.Text.Trim()
		//												, cboLoai.SelectedIndex });
		//	treeData.DataSource = arr;
		//	treeData.KeyFieldName = "ID";
		//	treeData.PreviewFieldName = "KnifeCode";
		//	treeData.ExpandAll();
		//}

		private void btnNewGroup_Click(object sender, EventArgs e)
		{
			frmKnifeGroupDetailH frm = new frmKnifeGroupDetailH();
			frm.Loai = cboLoai.SelectedIndex;
			if (frm.ShowDialog() == DialogResult.OK)
			{
				_curentNode = frm.CurentNode;
				loadTreeGroup();
			}
		}

		private void btnEditGroup_Click(object sender, EventArgs e)
		{
			try
			{
				if (treeDataGroup.AllNodesCount <= 0) return;
				int id = TextUtils.ToInt(treeDataGroup.FocusedNode.GetValue(colIDTreeGroup));
				if (id == 0) return;
				KnifeGroupModel model = (KnifeGroupModel)KnifeGroupBO.Instance.FindByPK(id);
				frmKnifeGroupDetailH frm = new frmKnifeGroupDetailH();
				frm.Loai = cboLoai.SelectedIndex;
				frm.KnifeGroup = model;
				if (frm.ShowDialog() == DialogResult.OK)
				{
					_curentNode = frm.CurentNode;
					loadTreeGroup();
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
				if (treeDataGroup.DataSource == null) return;
				int id = (int)treeDataGroup.Selection[0].GetValue(colIDTreeGroup);
				if (id == 0) return;

				if (KnifeDetailListBO.Instance.CheckExist("KnifeGroupID", id))
				{
					MessageBox.Show("Nhóm sản phẩm này đang được sử dụng nên không thể xóa được.", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}

				if (MessageBox.Show("Bạn có chắc muốn xóa nhóm [" + treeDataGroup.Selection[0].GetValue(colNameTreeGroup).ToString() + "] không?",
					  TextUtils.Caption, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
					return;
				KnifeGroupBO.Instance.Delete(id);
				loadTreeGroup();
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message, TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void treeDataGroup_FocusedNodeChanged(object sender, DevExpress.XtraTreeList.FocusedNodeChangedEventArgs e)
		{
			if (cboLoai.SelectedIndex == 0)
			{
				if (cbShowUnavailable.Checked)
				{
					LoadUnavailableData();
				}
				else
				{
					LoadData();
				}
				colATC.Visible = true;
				colATC.VisibleIndex = 4;
				colSTD.Visible = true;
				colSTD.VisibleIndex = 3;
				Visible1(true);
				Visible2(false);
			}
			else if (cboLoai.SelectedIndex == 3)
			{
				//Load Dụng cụ đo
				LoadToods();
				colATC.Visible = false;
				colSTD.Visible = false;
				Visible1(false);

				Visible2(true);
				return;
			}
			else
			{
				//LoadAll();
				//ToDo: Để sau khi click load theo nhóm
				try
				{
					//Load theo nhóm khi click
					LoadInfoSearch(0);
				}
				catch (Exception ex)
				{
					MessageBox.Show(ex.ToString());
				}
				Visible1(true);
				colATC.Visible = false;
				colSTD.Visible = false;
				Visible2(false);
			}


		}
		private void LoadInfoSearch(int type = 0)
		{
			if (treeDataGroup.AllNodesCount <= 0) return;
			if (type == 0)
			{
				_catID = TextUtils.ToInt(treeDataGroup.FocusedNode.GetValue(colIDTreeGroup));
			}
			else
			{
				_catID = 0;
			}
			DataTable arr = TextUtilsHP.LoadDataFromSP("spGetKnifeData"
								  , "A"
								  , new string[] { "@GroupID", "@TextFilter" }
								   , new object[] { _catID, txtFindDate.Text.Trim() }
					);
			treeDataKnife.DataSource = arr;
		}

		private void btnRestart_Click(object sender, EventArgs e)
		{
			LoadInfoSearch();
		}

		private void btnHistoryCheckOrderDetailH_Click(object sender, EventArgs e)
		{
			frmKnifeWorkingH frm = new frmKnifeWorkingH();
			frm.KnifeDetailListID = TextUtils.ToInt(treeDataKnife.FocusedNode.GetValue(colID));
			frm.ProductCode = TextUtils.ToString(treeDataKnife.FocusedNode.GetValue(colKnifeCode));
			frm.Loai = cboLoai.SelectedIndex;
			frm.Show();
		}
		private void btnImportExcel_Click(object sender, EventArgs e)
		{
			frmImportToolExcel frm = new frmImportToolExcel();
			frm.Show();
		}

		private void treeDataGroup_DoubleClick(object sender, EventArgs e)
		{

			btnEditGroup_Click(null, null);
		}

		private void treeDataKnife_DoubleClick(object sender, EventArgs e)
		{
			btnEditKnife_Click(null, null);
		}

		private void txtFindDate_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode != Keys.Enter) return;
			btnFindDate_Click(null, null);
		}

		private void btnPart_Click(object sender, EventArgs e)
		{

		}
	}
}
