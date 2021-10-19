using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using BMS.Business;
using IE.Model;
using DevExpress.Spreadsheet;
using DevExpress.Utils;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Web.UI.WebControls;

namespace BMS
{
    public partial class frmConfig : _Forms
    {
        public FontSize _FontSize;
        public frmConfig()
        {
            InitializeComponent();
        }

        #region Event
        private void frmConfig_Load(object sender, EventArgs e)
        {
        }
        #endregion

        private void numFontValueCD_ValueChanged(object sender, EventArgs e)
        {
            //_FontSize(numFontValueCD.Value, numFontTitleCD.Value, numFontValuePlan.Value, numFontLabelPlan.Value, numFontTitleAndon.Value, numLabelTakt.Value, numValueTakt.Value);
        }
    }
}
