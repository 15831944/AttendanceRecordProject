using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using Tools;
using System.Windows.Forms;
using Oracle.DataAccess.Client;
namespace AttendanceRecord.View
{
    public class V_AR_Time_Helper
    {
        #region 获取考勤时间
        public static DataTable getARTime(string _job_number,string _AR_Day) {
            string proceName = "PKG_ARTime.GET_JN_And_AR_Day";
            OracleParameter param_JN = new OracleParameter("v_job_number", OracleDbType.Varchar2, ParameterDirection.Input);
            OracleParameter param_AR_Day = new OracleParameter("v_ar_day", OracleDbType.Varchar2, ParameterDirection.Input);
            OracleParameter param_Cur_Result = new OracleParameter("v_cur_result", OracleDbType.RefCursor, ParameterDirection.Output);
            param_JN.Size = 20;
            param_AR_Day.Size = 20;
            //param_JN.Value = JN_Helper.getJN(_job_number);
            param_JN.Value = _job_number;
            param_AR_Day.Value = _AR_Day;
            OracleParameter[] parameters = new OracleParameter[3] { param_JN, param_AR_Day, param_Cur_Result };
            OracleHelper oH = OracleHelper.getBaseDao();
            DataTable dt = oH.getDT(proceName, parameters);
            return dt;
        }
        #endregion
        #region 获取考勤时间
        public static DataTable get_AR_Time(string _job_number, string _AR_Day)
        {
            string proceName = "PKG_TO_Export_From_A_R_Summary.GET_A_R_Time";
            OracleParameter param_JN = new OracleParameter("v_job_number", OracleDbType.Varchar2, ParameterDirection.Input);
            OracleParameter param_AR_Day = new OracleParameter("v_ar_day", OracleDbType.Varchar2, ParameterDirection.Input);
            OracleParameter param_Cur_Result = new OracleParameter("v_cur_result", OracleDbType.RefCursor, ParameterDirection.Output);
            param_JN.Size = 20;
            param_AR_Day.Size = 20;
            param_JN.Value = _job_number;
            param_AR_Day.Value = _AR_Day;
            OracleParameter[] parameters = new OracleParameter[3] { param_JN, param_AR_Day, param_Cur_Result };
            OracleHelper oH = OracleHelper.getBaseDao();
            DataTable dt = oH.getDT(proceName, parameters);
            return dt;
        }
        #endregion
        #region 获取考勤时间
        public static DataTable getARTimeByJN(string _job_number, string _AR_Day)
        {
            string proceName = "PKG_ARTime.GET_A_R_Time";
            OracleParameter param_JN = new OracleParameter("v_job_number", OracleDbType.Varchar2, ParameterDirection.Input);
            OracleParameter param_AR_Day = new OracleParameter("v_ar_day", OracleDbType.Varchar2, ParameterDirection.Input);
            OracleParameter param_Cur_Result = new OracleParameter("v_cur_result", OracleDbType.RefCursor, ParameterDirection.ReturnValue);
            param_JN.Size = 50;
            param_AR_Day.Size = 20;
            param_JN.Value = _job_number;
            param_AR_Day.Value = _AR_Day;
            OracleParameter[] parameters = new OracleParameter[3] { param_Cur_Result, param_JN, param_AR_Day  };
            OracleHelper oH = OracleHelper.getBaseDao();
            DataTable dt = oH.getDT(proceName, parameters);
            return dt;
        }
        #endregion
    }
}
