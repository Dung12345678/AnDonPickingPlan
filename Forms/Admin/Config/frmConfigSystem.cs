using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using BMS.Business;
using BMS.Utils;
using BMS.Model;

namespace BMS
{
    public partial class frmConfigSystem : _Forms
    {
        #region Variable
        ConfigSystemModel _model = new ConfigSystemModel();
        bool isAdd; 
        #endregion

        #region Constructor
        public frmConfigSystem()
        {
            InitializeComponent();
        }

        private void frmFixDirectory_Load(object sender, EventArgs e)
        {            
            LoadData();
        } 
        #endregion

        #region Methods
        public void LoadData()
        {            
            grdData.DataSource = ConfigSystemBO.Instance.FindAll();
            grvData.BestFitColumns();
        }

        private void SetInterface(bool isEdit)
        {
            txtDuongDan.Enabled = isEdit;
            txtKey.Enabled = isEdit;
            txtMoTa.Enabled = isEdit;
            grdData.Enabled = !isEdit;

            btnSave.Visible = isEdit;
            btnCancel.Visible = isEdit;

            btnNew.Visible = !isEdit;
            btnEdit.Visible = !isEdit;
            btnDelete.Visible = !isEdit;
        }

        private void ClearInterface()
        {
            txtDuongDan.Clear();
            txtKey.Clear();
            txtMoTa.Clear();
        }

        private bool checkValid()
        {
            if (txtKey.Text == "")
            {
                MessageBox.Show("Xin hãy điền tên.", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return false;
            }
            else
            {
                DataTable dt;
                if (_model.ID > 0)
                {
                    dt = TextUtils.Select("select KeyName from ConfigSystem where KeyName = '" + txtKey.Text.Trim() + "' and ID <> " + _model.ID);
                }
                else
                {
                    dt = TextUtils.Select("select KeyName from ConfigSystem where KeyName = '" + txtKey.Text.Trim() + "'");
                }
                if (dt != null)
                {
                    if (dt.Rows.Count > 0)
                    {
                        MessageBox.Show("Mã này đã tồn tại!", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                        return false;
                    }
                }
                
            }
            return true;
        }
        #endregion

        #region Event

        private void btnDong_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        
        private void btnAdd_Click(object sender, EventArgs e)
        {
            if (txtKey.Text == "" || txtDuongDan.Text == "")
            {
                MessageBox.Show("Cần nhập đầy đủ thông tin");
                return;
            }
            
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            try
            {

                if (MessageBox.Show("Bạn có chắc muốn xóa [" + grvData.GetFocusedRowCellValue(colKeyName).ToString() + "] không?",
                      TextUtils.Caption, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                    return;
                int id = TextUtils.ToInt(grvData.GetFocusedRowCellValue(colID).ToString());
                ConfigSystemBO.Instance.Delete(id);
                LoadData();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnHuy_Click(object sender, EventArgs e)
        {
            TextUtils.ClearTexbox(this);
            txtDuongDan.Clear();
            txtKey.Clear();
            txtMoTa.Clear();
        }        

        private void btnNew_Click(object sender, EventArgs e)
        {
            SetInterface(true);
            isAdd = true;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            try
            {
                if (checkValid())
                {
                    ProcessTransaction pt = new ProcessTransaction();
                    pt.OpenConnection();
                    pt.BeginTransaction();
                    if (!isAdd)
                    {
                        _model = (ConfigSystemModel)ConfigSystemBO.Instance.FindByPK(TextUtils.ToInt(grvData.GetFocusedRowCellValue(colID)));
                    }

                    _model.KeyName = txtKey.Text.Trim();
                    _model.KeyValue = txtDuongDan.Text.Trim();
                    _model.Description = txtMoTa.Text.Trim();

                    if (isAdd)
                    {
                        _model.CreatedDate = TextUtils.GetSystemDate();
                        _model.CreatedBy = Global.AppUserName;
                        _model.UpdatedDate = _model.CreatedDate;
                        _model.UpdatedBy = Global.AppUserName;
                        _model.ID = (int)pt.Insert(_model);
                    }
                    else
                    {
                        _model.UpdatedDate = TextUtils.GetSystemDate();
                        _model.UpdatedBy = Global.AppUserName;
                        pt.Update(_model);
                    }

                    pt.CommitTransaction();
                    pt.CloseConnection();
                    LoadData();
                    if (isAdd)
                    {
                        MessageBox.Show("Tạo mới thành công!");
                    }
                    else
                    {
                        MessageBox.Show("Cập nhật thành công!");
                    }
                    SetInterface(false);
                    ClearInterface();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            SetInterface(false);
            ClearInterface();
        }

        private void btnEdit_Click(object sender, EventArgs e)
        {
            if (!grvData.IsDataRow(grvData.FocusedRowHandle))
                return;
            SetInterface(true);
            
            int i = TextUtils.ToInt(grvData.GetFocusedRowCellValue(colID).ToString());
            _model = (ConfigSystemModel)ConfigSystemBO.Instance.FindByPK(i);

            txtKey.Text = grvData.GetFocusedRowCellValue(colKeyName).ToString();
            txtDuongDan.Text = grvData.GetFocusedRowCellValue(colKeyValue).ToString();
            txtMoTa.Text = grvData.GetFocusedRowCellValue(colDescription).ToString();
            isAdd = false;
        }     

        private void grvData_DoubleClick(object sender, EventArgs e)
        {
            btnEdit_Click(null, null);
        }
        #endregion
    }
}