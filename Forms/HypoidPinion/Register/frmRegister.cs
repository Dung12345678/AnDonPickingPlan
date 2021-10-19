using HP.Business;
using HP.Model;
using HP.Utils;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BMS

{
	public partial class frmRegister : _Forms
	{
		DataTable _dtTools;
		public frmRegister()
		{
			InitializeComponent();
		}

		private void frmProductdetail_Load(object sender, EventArgs e)
		{
			cboKnifeDetailList.SelectedIndex = 0;
			loadGoods();
		}

		private void loadGoods()
		{
			string sql = "Select * from [HypoidPinion].[dbo].[Goods]";
			DataTable dt = TextUtilsHP.Select(sql);
			cboMaHang.Properties.DisplayMember = "Name";
			cboMaHang.Properties.ValueMember = "ID";
			cboMaHang.Properties.DataSource = dt;
		}
		private void loadTool()
		{
			_dtTools = TextUtilsHP.GetDataTableFromSP("spLoadRegister", new string[] { "@type", "@GoodsID" }, new object[] { cboKnifeDetailList.SelectedIndex, cboMaHang.EditValue });
			grdData.DataSource = _dtTools;
		}

		private void cboMaHang_EditValueChanged(object sender, EventArgs e)
		{
			loadTool();
		}

		private void btnSave_Click(object sender, EventArgs e)
		{
			try
			{
				grvData.Focus();
				cboMaHang.Focus();
				if (cboMaHang.Text.Trim() == "") return;
				DialogResult rs = MessageBox.Show("Bạn chắc chắn muốn cất", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
				if (rs == DialogResult.No) return;

				for (int i = 0; i < grvData.RowCount; i++)
				{
					int registerID = 0;
					int check = TextUtilsHP.ToInt(grvData.GetRowCellValue(i, colOk));
					registerID = TextUtilsHP.ToInt(grvData.GetRowCellValue(i, colRegisterID));
					Expression exp = new Expression("GoodsID", cboMaHang.EditValue);
					Expression exp1 = new Expression("ToolsID", TextUtilsHP.ToInt(grvData.GetRowCellValue(i, colKnifeDetailListID)));
					//Insert vào bảng Dao jig khuôn
					if (cboKnifeDetailList.SelectedIndex <= 2)
					{
						RegisterModel registerModel = new RegisterModel();
						//Insert vào lịch sử 
						RegisterHistoryModel registerHistoryModel = new RegisterHistoryModel();
						registerHistoryModel.GoodsID = TextUtilsHP.ToInt(cboMaHang.EditValue);
						registerHistoryModel.ToolsID = TextUtilsHP.ToInt(grvData.GetRowCellValue(i, colKnifeDetailListID));
						registerHistoryModel.IsCheck = TextUtilsHP.ToBoolean(grvData.GetRowCellValue(i, colOk));
						registerHistoryModel.CreateDate = DateTime.Now;
						RegisterHistoryBO.Instance.Insert(registerHistoryModel);

						if (check <= 0)
						{
							//registerModel = (RegisterModel)RegisterBO.Instance.FindByPK(registerID);
							//Nếu không check xóa các thông tin trong bảng Register
							RegisterBO.Instance.DeleteByExpression(exp.And(exp1));
							continue;
						}
						//Nếu check thì tìm thông tin trong bảng có update k thì insert
						ArrayList array = RegisterBO.Instance.FindByExpression(exp.And(exp1));
						if (array.Count > 0)
							registerModel = (RegisterModel)array[0];

						registerModel.GoodsID = TextUtilsHP.ToInt(cboMaHang.EditValue);
						registerModel.ToolsID = TextUtilsHP.ToInt(grvData.GetRowCellValue(i, colKnifeDetailListID));

						if (array.Count > 0)
						{
							RegisterBO.Instance.Update(registerModel);
						}
						else
						{
							RegisterBO.Instance.Insert(registerModel);
						}

						//Update vào bảng Dao 
						if (cboKnifeDetailList.SelectedIndex == 0)
						{
							//KnifeDetailListModel knifeDetailListModel = (KnifeDetailListModel)KnifeDetailListBO.Instance.FindByPK(registerModel.ToolsID);
							//knifeDetailListModel.CurrentSTD += TextUtilsHP.ToInt(TextUtilsHP.ExcuteScalar($"SELECT COUNT(OrderMachining) FROM [HypoidPinion].[dbo].[ProductKnife] WHERE GoodsCode='{cboMaHang.Text.Trim()}'"));
							//if (knifeDetailListModel.CurrentSTD == 0) continue;
							//KnifeDetailListBO.Instance.Update(knifeDetailListModel);
						}
					}
					else
					{
						if (check <= 0)
						{
							//registerModel = (RegisterModel)RegisterBO.Instance.FindByPK(registerID);
							//Nếu không check xóa các thông tin trong bảng Register
							GoodsRegisterBO.Instance.DeleteByExpression(exp.And(exp1));
							continue;
						}
						GoodsRegisterModel goodsRegisterModel = new GoodsRegisterModel();
						//Nếu check thì tìm thông tin trong bảng có update k thì insert
						ArrayList array = GoodsRegisterBO.Instance.FindByExpression(exp.And(exp1));
						if (array.Count > 0)
							goodsRegisterModel = (GoodsRegisterModel)array[0];

						goodsRegisterModel.GoodsID = TextUtilsHP.ToInt(cboMaHang.EditValue);
						goodsRegisterModel.ToolsID = TextUtilsHP.ToInt(grvData.GetRowCellValue(i, colKnifeDetailListID));

						if (array.Count > 0)
						{
							GoodsRegisterBO.Instance.Update(goodsRegisterModel);
						}
						else
						{
							GoodsRegisterBO.Instance.Insert(goodsRegisterModel);
						}
					}
				}


				MessageBox.Show("Thêm thành công", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);


				//DataTable dt1 = new DataTable();
				//dt1 = grvData.GetDataRow(0).Table;

			}
			catch
			{

			}
		}

		private void cboKnifeDetailList_SelectedIndexChanged(object sender, EventArgs e)
		{
			loadTool();
			if (cboKnifeDetailList.SelectedIndex == 3)
			{
				colKnifeName.Visible = false;
				colKnifeCode.Visible = false;
				colCodeTool.Visible = true;
				colNameTool.Visible = true;
			}
			else
			{
				colKnifeName.Visible = true;
				colKnifeCode.Visible = true;
				colCodeTool.Visible = false;
				colNameTool.Visible = false;
			}	
		}

		private void btnImportExcel_Click(object sender, EventArgs e)
		{
			frmImportGoodsRegisterExcel frm = new frmImportGoodsRegisterExcel();
			frm.Show();
		}

		private void btnRestart_Click(object sender, EventArgs e)
		{
			loadGoods();
		}
	}
}
