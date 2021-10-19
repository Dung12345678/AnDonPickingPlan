using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using BMS.Business;
using BMS.Utils;
using BMS.Model;
using DevExpress.XtraGrid.Localization;

namespace BMS
{
    public partial class frmHistory : _Forms
    {
        public int ID = 0;
        public string Code = "";
        //public ModulesModel Module = new ModulesModel();
        public frmHistory()
        {
            InitializeComponent();
        }

        private void frmModuleHistory_Load(object sender, EventArgs e)
        {
            GridLocalizer.Active = new MyGridLocalizer();

            DataTable dt = new DataTable();
            dt = TextUtils.Select("SELECT * FROM ActivityLog with(nolock) order by DateProcess desc");
            grdData.DataSource = null;
            grdData.DataSource = dt;
        }

        private void grvData_DoubleClick(object sender, EventArgs e)
        {
            DevExpress.XtraGrid.Views.Grid.GridView sndr =
                   sender as DevExpress.XtraGrid.Views.Grid.GridView;
            DevExpress.Utils.DXMouseEventArgs dxMouseEventArgs =
                e as DevExpress.Utils.DXMouseEventArgs;
            DevExpress.XtraGrid.Views.Grid.ViewInfo.GridHitInfo hitInfo =
               sndr.CalcHitInfo(dxMouseEventArgs.Location);

            if (hitInfo.InColumn)
            {
                grvData.ShowCustomFilterDialog(hitInfo.Column);
            }
        }
    }
}
