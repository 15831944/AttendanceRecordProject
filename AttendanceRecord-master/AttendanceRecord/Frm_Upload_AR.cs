using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Tools;

namespace AttendanceRecord
{
    public partial class Frm_Upload_AR : Form
    {
        string defaultDir = Environment.CurrentDirectory + @"\考勤记录";
        public static string _action = "upload_AR";
        string xlsFilePath = string.Empty;
        public Frm_Upload_AR()
        {
            InitializeComponent();
        }

        private void btnImportEmpsInfo_Click(object sender, EventArgs e)
        {
            FTPHelper ftpHelper = new FTPHelper();
            xlsFilePath = FileNameDialog.getSelectedFilePathWithDefaultDir("请选择考勤记录：", "*.xls|*.xls", defaultDir);
            tb.Text = xlsFilePath;
            string dir = DirectoryHelper.getDirOfFile(xlsFilePath);
            if (string.IsNullOrEmpty(dir))
            {
                return;
            }
            List<string> xlsFilePathList = DirectoryHelper.getXlsFileUnderThePrescribedDir(dir);
            for (int i = 0; i <= xlsFilePathList.Count - 1; i++)
            {
                //上传文件.
                ftpHelper.UpLoadFile(xlsFilePathList[i], ftpHelper.FtpURI + DirectoryHelper.getFileName(xlsFilePathList[i]));
            }
            //隐藏
            this.MdiParent.WindowState = FormWindowState.Minimized;
            this.MdiParent.ShowInTaskbar = false;
            //启动定时器


        }
        /// <summary>
        /// 从服务器读取处理进度。
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void timerReadProcessFromSvr_Tick(object sender, EventArgs e)
        {

        }
    }
}
