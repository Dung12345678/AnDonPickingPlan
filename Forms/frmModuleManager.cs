using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using BMS.Utils;
using BMS.Model;
using BMS.Business;
using System.Collections;
using DevExpress.XtraTreeList.Nodes;
using DevExpress.XtraTreeList;
using System.Reflection;
using System.Threading;
using DevExpress.XtraGrid.Views.Grid;
using System.IO;
using System.Diagnostics;
using DevExpress.XtraGrid.Localization;
using iTextSharp;
using iTextSharp.text.pdf;
using DevExpress.Utils;
using IOEx;

namespace BMS
{
    public partial class frmModuleManager : _Forms
    {
        #region Variables
        #endregion

        #region Constructors
        public frmModuleManager()
        {
            InitializeComponent();
        }

        private void frmModuleManager_Load(object sender, EventArgs e)
        {
           
        }
        #endregion

    }    
}
