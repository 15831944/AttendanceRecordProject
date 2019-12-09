using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Tools;
using System.Data;
using Oracle.DataAccess.Client;
namespace AttendanceRecord.View
{
    /// <summary>
    /// 未打卡视图。
    /// 1.计算当天未打卡次数。
    /// 2.给定月份计算，未打卡次数。
    /// 
    /// </summary>
    public class V_Not_Finger_Print
    {
        /// <summary>
        /// 返回给定日期未打开的次数。
        /// </summary>
        /// <param name="_year_and_month"></param>
        /// <returns></returns>
        public static int get_not_FingerPrint_Times(string _job_Number,string _year_and_month) {
            int count = 0;
            string procedure_name = "Get_Not_FingerPrint_Times";
            OracleParameter param_JN = new OracleParameter("v_Job_Number", OracleDbType.Varchar2, ParameterDirection.Input);
            OracleParameter param_year_and_month = new OracleParameter("v_year_and_month", OracleDbType.Varchar2, ParameterDirection.Input);
            param_JN.Value = _job_Number;
            param_year_and_month.Value = _year_and_month;
            OracleParameter param_not_FingerPrint_Times = new OracleParameter("v_not_FingerPrint_times",OracleDbType.Int32,ParameterDirection.Output);
            OracleParameter[] parameters = new OracleParameter[3] { param_JN,param_year_and_month,param_not_FingerPrint_Times};
            OracleHelper oH = OracleHelper.getBaseDao();
            oH.ExecuteNonQuery(procedure_name, parameters);
            count = int.Parse(parameters[2].Value.ToString());
            return count;
        }
    }
}
