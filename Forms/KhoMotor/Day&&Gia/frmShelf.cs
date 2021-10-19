using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.CodeDom.Compiler;
using Forms;
using System.Diagnostics;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.Utils;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;
using System.Threading;
using System.Threading.Tasks;
using System.IO;
using ST.Model;
using ST.Business;

namespace BMS
{
	public partial class frmShelf : _Forms
	{
		DataTable _dtList = new DataTable();
		DataTable _dtB = new DataTable();
		DataTable _dtName = new DataTable();
		public SequenceModel name = new SequenceModel();
		public frmShelf()
		{
			InitializeComponent();
		}

		private void frmListTool_Load(object sender, EventArgs e)
		{
			loadData();
		}

		private void loadData()
		{
			DataTable dt = new DataTable();
			dt = TextUtilsStock.Select("SELECT * FROM dbo.Sequence");
			treeData.DataSource = dt;
		}
		DataTable dttt = new DataTable();
		void loadGrdData()
		{
			int GroupId = TextUtilsStock.ToInt(treeData.FocusedNode.GetValue(colIDTree));
			DataTable dt = new DataTable();
			dt = TextUtilsStock.LoadDataFromSP("spLoadShelfQL", "A", new string[] { "@GroupId" }, new object[] { GroupId });
			dttt = dt;
			DataColumn data = new DataColumn("Location", typeof(Byte[]));
			dt.Columns.Add(data);
			grdData.DataSource = dt;

		}
		/// <summary>
		/// load dữ liệu vào bảng grdData khi thay đổi tại bảng treeData
		/// </summary>
		//public string GroupName;
		private void treeData_FocusedNodeChanged(object sender, DevExpress.XtraTreeList.FocusedNodeChangedEventArgs e)
		{
			int GroupId = TextUtilsStock.ToInt(treeData.FocusedNode.GetValue(colIDTree));
			if (GroupId > 0)
			{
				loadGrdData();
			}
		}
		/// <summary>
		/// click button tạo nhóm sản phẩm
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnNewGroup_Click(object sender, EventArgs e)
		{
			frmSequence frm = new frmSequence();
			if (frm.ShowDialog() == DialogResult.OK)
			{
				loadData();
			}
		}
		/// <summary>
		/// hàm sửa lại dữ liệu Group
		/// </summary>
		public static int EditClick = 0;
		private void editDataGroup()
		{
			//EditClick = 1;
			int id = TextUtilsStock.ToInt(treeData.FocusedNode.GetValue(colIDTree));
			SequenceModel model = (SequenceModel)SequenceBO.Instance.FindByPK(id);
			frmSequence frm = new frmSequence();
			frm._SequenceModel = model;
			if (frm.ShowDialog() == DialogResult.OK)
			{
				loadData();
			}

		}
		/// <summary>
		/// dùng button Edit Data Group
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnEditGroup_Click(object sender, EventArgs e)
		{
			if (treeData.AllNodesCount < 0) return;
			editDataGroup();
		}
		/// <summary>
		/// click double to edit data group
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void treeData_DoubleClick(object sender, EventArgs e)
		{
			editDataGroup();
		}

		private void btnDeleteGroup_Click(object sender, EventArgs e)
		{
			if (treeData.AllNodesCount < 0) return;
			int groupID = TextUtilsStock.ToInt(treeData.FocusedNode.GetValue(colIDTree));
			string groupName = TextUtilsStock.ToString(treeData.FocusedNode.GetValue(colNameTree));
			if (groupID == 0) return;

			if (ShelfBO.Instance.CheckExist("SequenceID", groupID))
			{
				MessageBox.Show(string.Format("Nhóm thiết bị đang chứa thiết bị, không thể xóa nhóm thiết bị [{0}] khỏi danh sách!!", groupName), TextUtilsStock.Caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			if (MessageBox.Show(string.Format("Bạn có muốn xóa nhóm [{0}] hay không ?", groupName), TextUtilsStock.Caption, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
			{
				SequenceBO.Instance.Delete(groupID);
				treeData.DeleteNode(treeData.FocusedNode);
			}
		}

		private void deleteToolStripMenuItem_Click(object sender, EventArgs e)
		{
			btnDeleteGroup_Click(null, null);
		}

		//--------------- code phần sản phẩm, thiết bị ------------------------------------
		/// <summary>
		/// creat tool
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnAdd_Click(object sender, EventArgs e)
		{
			int ID = TextUtilsStock.ToInt(treeData.FocusedNode.GetValue(colIDTree));
			string groupName = TextUtilsStock.ToString(treeData.FocusedNode.GetValue(colNameTree));
			if (ID == 0) return;
			ShelfModel model = new ShelfModel();
			model.SequenceID = ID;
			frmShelfDetail frm = new frmShelfDetail();
			frm.oShelfModel = model;
			if (frm.ShowDialog() == DialogResult.OK)
			{
				loadGrdData();
			}
		}
		/// <summary>
		/// fix tool
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnEdit_Click(object sender, EventArgs e)
		{
			if (grvData.RowCount < 0) return;
			editDataProduct();
		}
		/// <summary>
		/// void edit data in productRTC
		/// </summary>
		public static int editGrv = 0;
		private void editDataProduct()
		{
			editGrv = 1;
			int ID = TextUtilsStock.ToInt(grvData.GetFocusedRowCellValue(colID));
			if (ID == 0) return;
			ShelfModel model = (ShelfModel)ShelfBO.Instance.FindByPK(ID);
			frmShelfDetail frm = new frmShelfDetail();
			frm.oShelfModel = model;
			if (frm.ShowDialog() == DialogResult.OK)
			{
				loadGrdData();
			}
		}
		/// <summary>
		/// delete sản phẩm khỏi danh sách
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnDel_Click(object sender, EventArgs e)
		{
			if (grvData.RowCount < 0) return;
			int ID = TextUtilsStock.ToInt(grvData.GetFocusedRowCellValue(colID));
			string code = TextUtilsStock.ToString(grvData.GetFocusedRowCellValue(colShelfCode));
			string name = TextUtilsStock.ToString(grvData.GetFocusedRowCellValue(colShelfName));
			if (ID == 0) return;

			if (MessageBox.Show(string.Format("Bạn có chắc chắn muốn xóa mã giá đỡ [{0}] : [{1}] không?", code, name), TextUtilsStock.Caption, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
			{
				ShelfBO.Instance.Delete(ID);
				grvData.DeleteSelectedRows();
			}
		}
		/// <summary>
		/// event editData by doubleClick
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void grdData_DoubleClick(object sender, EventArgs e)
		{
			int id = TextUtilsStock.ToInt(grvData.GetFocusedRowCellValue(colID));
			ShelfModel model = (ShelfModel)ShelfBO.Instance.FindByPK(id);
			if (TextUtilsStock.ToBoolean(grvData.GetFocusedRowCellValue(colShelfCode)) == false)
			{
				frmShelfDetail frm = new frmShelfDetail();
				frm.oShelfModel = model;
				if (frm.ShowDialog() == DialogResult.OK)
				{
					loadData();
				}
			}
		}
		private void btn_Click(object sender, EventArgs e)
		{
		}
	}
}


