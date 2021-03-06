using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using BMS.Utils;
using System.IO;

namespace BMS
{
    public partial class frmAppSettings : _Forms
    {
        public frmAppSettings()
        {
            InitializeComponent();
        }

        private void frmAppSettings_Load(object sender, EventArgs e)
        {
            try
            {
                cboCOMPort.SelectedIndex = 0;
                string[] settings = new string[1];
                FileStream fs = null;
                try
                {
                    fs = File.Open(Application.StartupPath + @"\settings.ini", FileMode.Open, FileAccess.ReadWrite, FileShare.None);
                }
                catch
                {
                    try
                    {
                        fs = File.Create(Application.StartupPath + @"\settings.ini");
                        StreamWriter sw = new StreamWriter(fs);
                        sw.WriteLine("[settings]");
                        sw.WriteLine("COMPort=COM7");
                        sw.Flush();
                        sw.Close();
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                }
                finally
                {
                    fs.Close();
                }
                settings = File.ReadAllLines(Application.StartupPath + @"\settings.ini");
                foreach (string st in settings)
                {
                    if (st.StartsWith("COMPort"))
                    {
                        cboCOMPort.SelectedItem = st.Split('=')[1];
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            try
            {
                FileStream fs = File.Open(Application.StartupPath + @"\settings.ini", FileMode.Truncate, FileAccess.ReadWrite, FileShare.None);
                StreamWriter sw = new StreamWriter(fs);
                sw.WriteLine("[settings]");
                sw.WriteLine("COMPort=" + cboCOMPort.SelectedItem);
                sw.Flush();
                sw.Close();
                MessageBox.Show("Thiết lập tùy chọn xong!", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Close();
            }
            catch
            {
                MessageBox.Show("Có lỗi xảy ra trong quá trình lưu tùy chọn.\nXin vui lòng thử lại sau.", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            this.Close();
        }
    }
}