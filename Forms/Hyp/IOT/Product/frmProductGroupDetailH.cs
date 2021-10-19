using DevExpress.Utils;
using IE.Business;
using IE.Model;
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
    public partial class frmProductGroupDetailH : _Forms
    {
        public ProductGroupModel ProductGroup = new ProductGroupModel();
        ArrayList arrStep = new ArrayList();

        public int CurentNode { get; set; }
        string _Name = "";
        int _STT = 0
            ;
        string _Des = "";

        public frmProductGroupDetailH()
        {
            InitializeComponent();
        }

        private void frmProductGroup_Load(object sender, EventArgs e)
        {
            txtCode.Text = ProductGroup.ProductGroupCode;
            txtName.Text = ProductGroup.ProductGroupName;

            loadData();

            DataTable dt = LibIE.Select("SELECT * FROM ProductGroup");
            cbCopy.Properties.DataSource = dt;
            cbCopy.Properties.DisplayMember = "ProductGroupName";
            cbCopy.Properties.ValueMember = "ID";
        }

        #region Methods
        private void loadData()
        {
            try
            {
                arrStep = WorkingStepBO.Instance.GetListObject("spGetWorkingStep_ByGroupID", new string[] { "@GroupID" }, new object[] { ProductGroup.ID });
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
            frmWorkingStepDetailH frm = new frmWorkingStepDetailH();
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
                int id = TextUtils.ToInt(grvData.GetFocusedRowCellValue(colID));
                int sortOrder = TextUtils.ToInt(grvData.GetFocusedRowCellValue(colSortOrder));
                WorkingStepModel model = new WorkingStepModel();
                var dr = grvData.GetRow(grvData.FocusedRowHandle);
                model = (WorkingStepModel)dr;
                frmWorkingStepDetailH frm = new frmWorkingStepDetailH();
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

            int strID = TextUtils.ToInt(grvData.GetFocusedRowCellValue("ID"));

            string strName = TextUtils.ToString(grvData.GetFocusedRowCellValue("SortOrder"));

            if (WorkingBO.Instance.CheckExist("WorkingStepID", strID))
            {
                MessageBox.Show("Công đoạn này đang được sử dụng bên Mục kiểm tra nên không thể xóa được!", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            if (MessageBox.Show(String.Format("Bạn có chắc muốn xóa công đoạn [{0}] không?", strName), TextUtils.Caption, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                try
                {
                    if (strID > 0)
                    {
                        LibIE.ExcuteProcedure("spUpdateProductStep_ByWorkingStep"
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
                    ProductGroup.ProductGroupCode = txtCode.Text.Trim();
                    ProductGroup.ProductGroupName = txtName.Text.Trim();
                    if (ProductGroup.ID <= 0)
                    {
                        ProductGroup.ID = (int)ProductGroupBO.Instance.Insert(ProductGroup);
                    }
                    else
                    {
                        ProductGroupBO.Instance.Update(ProductGroup);
                    }
                    CurentNode = ProductGroup.ID;
                    for (int i = 0; i < arrStep.Count; i++)
                    {
                        _STT = Lib.ToInt(grvData.GetRowCellValue(i,colSortOrder));
                        _Name = Lib.ToString(grvData.GetRowCellValue(i,colName));
                        _Des = Lib.ToString(grvData.GetRowCellValue(i,colDes));
                        WorkingStepModel step = (WorkingStepModel)arrStep[i];
                        int type = 0;
                        if (step.ID <= 0)
                        {
                            step.ProductGroupID = ProductGroup.ID;
                            step.ID = (int)WorkingStepBO.Instance.Insert(step);
                            type = 1;
                            //Thêm step vào các sản phẩm thuộc nhóm này
                        }
                        else
                        {
                            //string sql= $"SELECT TOP 1 ws.ID FROM dbo.WorkingStep ws JOIN dbo.ProductStep ps ON ws.ID = ps.WorkingStepID WHERE WorkingStepID = {step.ID} and ws.WorkingStepCode = N'{_Name}'  and ws.Description = N'{_Des}' and ws.SortOrder ='{_STT}'";
                            //int id = Lib.ToInt(LibIE.ExcuteScalar(sql));
                            //if (id == 0)
                            //{
                                WorkingStepBO.Instance.Update(step);
                                type = 2;
                                // type = 1;
                                //Cập nhật vào các sản phẩm thuộc nhóm này
                           //}
                        }

                        LibIE.ExcuteProcedure("spUpdateProductStep_ByWorkingStep"
                            , new string[] { "@WorkingStepID", "@Type" }
                            , new object[] { step.ID, type });

                    }

                    DialogResult = DialogResult.OK;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                model.WorkingStepCode = Lib.ToString(grvData.GetRowCellValue(i, colName));
                model.SortOrder = Lib.ToInt(grvData.GetRowCellValue(i, colSortOrder));
                model.Description = Lib.ToString(grvData.GetRowCellValue(i, colDes));
                model.CreatedBy = "Admin";
                model.CreatedDate = DateTime.Now;
                model.UpdatedBy = "Admin";
                model.UpdatedDate = DateTime.Now;
                model.ProductGroupID = Lib.ToInt(cbCopy.EditValue);
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
