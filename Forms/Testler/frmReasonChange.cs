using BMS;
using IE.Business;
using IE.Model;
using System;
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
    public partial class frmReasonChange : _Forms
    {
        public string Reason;
        public frmReasonChange()
        {
            InitializeComponent();
        }       
        private void frmReasonChange_Load(object sender, EventArgs e)
        {
            txtText.Text = Reason;
        }
        private void txtText_KeyDown(object sender, KeyEventArgs e)
        {
                             
        }
        private void btnAccept_Click(object sender, EventArgs e)
        {
            Reason = this.txtText.Text.Trim();
            DialogResult = DialogResult.OK;
        }

        private void frmReasonChange_FormClosed(object sender, FormClosedEventArgs e)
        {
            //Reason = this.txtText.Text.Trim();
            //DialogResult = DialogResult.OK;
            this.Close();
        }
    }
}
