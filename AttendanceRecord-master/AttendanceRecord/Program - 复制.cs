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
    public static class ProgramBackup
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
        static void MainBackup()
        {
            XmlFlexflow.configFilePath = Application.StartupPath + "\\flexflow.cfg";
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            string ftpIPADDR = XmlFlexflow.ReadXmlNodeValue("FTP_IPADDR");
            //先测试是否可以ping通
            if (!ConnectByPing.pingTheAddress(ftpIPADDR))
            {
                if (DialogResult.No == MessageBox.Show("未能与版本服务器取得联系,是否继续？", "提示: ", MessageBoxButtons.YesNo, MessageBoxIcon.Information))
                {
                    return;
                }
                //继续。
                doNextStep();
                return;
            }
            string releaseVersion = string.Empty;
            //检查软件版本
            MSG msg = CheckAppVerionByFTP.checkAppVersionAndOutReleaseVersion(ftpIPADDR, out releaseVersion);
            if (!msg.Flag)
            {
                //MessageBox.Show(msg.Msg, "提示：", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                return;
            }
            doNextStep();
        }
        /// <summary>
        /// 
        /// </summary>
        static void doNextStep()
        {
            #region 数据库联接测试。
            string host_Name = XmlFlexflow.ReadXmlNodeValue("SERVER_NAME");
            string service_Name = XmlFlexflow.ReadXmlNodeValue("SERVICE_NAME");
            string user_Id = XmlFlexflow.ReadXmlNodeValue("USER_ID");
            string password = XmlFlexflow.ReadXmlNodeValue("PASSWORD");

            string mes_host_Name = XmlFlexflow.ReadXmlNodeValue("MES_SERVER_NAME");
            string mes_db_Name = XmlFlexflow.ReadXmlNodeValue("MES_DATABASE_NAME");
            string mes_user_Id = XmlFlexflow.ReadXmlNodeValue("MES_USER_ID");
            string mes_password = XmlFlexflow.ReadXmlNodeValue("MES_PASSWORD");
            //先测试是否可以ping通
            if (!ConnectByPing.pingTheAddress(host_Name))
            {
                MessageBox.Show("与" + host_Name + " 连接失败!", "提示: ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            #endregion
            //再这个启动类里，对对象进行初始化。
            OracleDaoHelper daoHelper = new OracleDaoHelper(host_Name, service_Name, user_Id, password);
            SqlDaoHelper sqlHelper = new SqlDaoHelper(mes_host_Name, mes_db_Name, mes_user_Id, mes_password);
            SqlConnection sqlConn = new SqlConnection(SqlDaoHelper.conn_str);
            try
            {
                sqlConn.Open();
                sqlConn.Close();
                sqlConn.Dispose();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                MessageBox.Show("基于MES_制卡系统中的所属部门，组将无法获取");
            }
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
