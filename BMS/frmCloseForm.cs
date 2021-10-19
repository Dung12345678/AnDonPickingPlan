using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace BMS
{
    public partial class frmCloseForm : Form
    {
        public frmCloseForm()
        {
            InitializeComponent();            
        }

        private void frmCloseForm_Load(object sender, EventArgs e)
        {
            Application.Exit();
        }
    }
}
