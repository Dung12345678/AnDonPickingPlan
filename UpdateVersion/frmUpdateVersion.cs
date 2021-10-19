using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
namespace BMS.UP
{
    public partial class frmUpdateVersion : Form
    {
        string _pathFileConfigUpdate = Path.Combine(Application.StartupPath, "ConfigUpdate.txt");
        string _extractPath = Path.GetFullPath(Application.StartupPath);
        string _pathFolderUpdate = "";
        string _pathUpdateServer = "";
        string _appName = "";
        string _pathFileVersion = "";
        bool _isErrorExtract = false;

        public frmUpdateVersion()
        {
            InitializeComponent();
        }
        void TurnOffApp(string appName)
        {
            Process[] arr = Process.GetProcessesByName(appName);
            for (int i = 0; i < arr.Length; i++)
            {
                arr[i].Kill();
            }
            Thread.Sleep(1000);
        }

        void updateVersion(string PathFileVersion)
        {
            try
            {                
                string[] listFileSv = DocUtils.GetFilesList(@_pathUpdateServer);
                List<string> lst = listFileSv.ToList();
                lst = lst.Where(o => o.Contains(".zip")).ToList();
                int newVersion = lst.Max(o => TextUtils.ToInt(Path.GetFileNameWithoutExtension(o)));
                string fileName = newVersion + ".zip";

                if (File.Exists(Path.Combine(_pathFolderUpdate, fileName)))
                {
                    File.Delete(Path.Combine(_pathFolderUpdate, fileName));
                }
                DocUtils.DownloadFile(_pathFolderUpdate, fileName, Path.Combine(_pathUpdateServer, fileName));
                if (File.Exists(Path.Combine(_pathFolderUpdate, fileName)))
                {
                    if (!_extractPath.EndsWith(Path.DirectorySeparatorChar.ToString()))
                        _extractPath += Path.DirectorySeparatorChar;
                    using (ZipArchive archive = ZipFile.OpenRead(Path.Combine(_pathFolderUpdate, fileName)))
                    {                            
                        foreach (ZipArchiveEntry entry in archive.Entries)
                        {
                            try
                            {
                                string destinationPath = Path.GetFullPath(Path.Combine(_extractPath, entry.FullName));
                                if (destinationPath.StartsWith(_extractPath, StringComparison.Ordinal))
                                    entry.ExtractToFile(destinationPath, true);                                   
                            }
                            catch (Exception e)
                            {
                                MessageBox.Show(e.ToString());
                                _isErrorExtract = true;
                                break;
                            }   
                        }
                        if (!_isErrorExtract)
                        {
                            File.WriteAllText(PathFileVersion, TextUtils.ToString(newVersion));
                        }
                    }
                }
                
            }
            catch(Exception ex) {

            }
        }
        private void Form1_Load(object sender, EventArgs e)
        {            
            string[] arrSv = File.ReadAllLines(Path.Combine(Application.StartupPath, "ftpServer.txt"));
            if (arrSv.Length < 3) MessageBox.Show("Lỗi file ftpServer.txt. Hãy kiểm tra lại! ");
           
            DocUtils.InitFTPQLSX(arrSv[0].Trim(), arrSv[1].Trim(), arrSv[2].Trim());
            string[] lines = File.ReadAllLines(_pathFileConfigUpdate);
            string[] stringSeparators = new string[] { "||" };
            string[] arr = lines[1].Split(stringSeparators, 4, StringSplitOptions.RemoveEmptyEntries);
            _appName = arr[0].Trim();
            _pathFolderUpdate = Path.Combine(Application.StartupPath, arr[1].Trim());
            _pathUpdateServer = arr[2].Trim();
            _pathFileVersion = Path.Combine(Application.StartupPath, arr[3].Trim());
            if (backgroundWorker1.IsBusy)
            {
                backgroundWorker1.CancelAsync();
            }
            else
            {
                //Tắt các ứng dụng RTC
                try
                {
                    TurnOffApp(_appName);
                    backgroundWorker1.RunWorkerAsync();
                }
                catch (Exception ex)
                {
                    Application.Exit();
                }
            }
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                updateVersion(_pathFileVersion);                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (!_isErrorExtract)
            {
                Process.Start(Path.Combine(Application.StartupPath, _appName + ".exe"));
                Application.Exit();
            }
        }
    }
}
