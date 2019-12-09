using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using AttendanceRecord.Entities;
using Tools;
using AttendanceRecord.Helper;
using System.Data.SqlClient;
namespace AttendanceRecord
{
    public static class Program
    {
        #region Version Info
        //=====================================================================
        // Project Name        :    BaseDao  
        // Project Description : 
        // Class Name          :    Class1
        // File Name           :    Class1
        // Namespace           :    BaseDao 
        // Class Version       :    v1.0.0.0
        // Class Description   : 
        // CLR                 :    4.0.30319.42000  
        // Author              :    董   魁  (ccie20079@126.com)
        // Addr                :    中国  陕西 咸阳    
        // Create Time         :    2019-10-22 14:57:19
        // Modifier:     
        // Update Time         :    2019-10-22 14:57:19
        //======================================================================
        // Copyright © DGCZ  2019 . All rights reserved.
        // =====================================================================
        #endregion
        public static User_Info _userInfo;
        public static bool flag_open_mesSqlConn = false;
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            XmlFlexflow.configFilePath = Application.StartupPath + "\\flexflow.cfg";
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            bool resultFlag = DownloadTheLatestApp.downloadTheLatestVersionAndInitConnStr();
            if (!resultFlag) return;
            //设置连接数据库字符串的值.
            SetTheValueOfTheDatabaseConnStr.setTheValueOfTheConnStr();
            doNext();
        }
        /// <summary>
        /// 
        /// </summary>
        static void doNext()
        {
            FormLogin frmLogin = new FormLogin();
            frmLogin.ShowDialog();
            if (DialogResult.OK != frmLogin.DialogResult)
            {
                //结束程序
                return;
            }
            FrmMainOfAttendanceRecord frmMainOfAttendanceRecord = new FrmMainOfAttendanceRecord();
            Application.Run(frmMainOfAttendanceRecord);
        }
    }
}
