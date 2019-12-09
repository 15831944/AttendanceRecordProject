using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Oracle.DataAccess.Client;
using System.Data;
using Tools;
namespace AttendanceRecord.View
{
    /// <summary>
    /// 此类用来计算实际出勤天数。
    /// </summary>
    public class V_ActualAttendanceDays
    {
        private string _job_Number;
        private string _year_and_month;

        public V_ActualAttendanceDays(string _job_Number, string _year_and_month)
        {
            this._job_Number = _job_Number;
            this._year_and_month = _year_and_month;
        }
        #region 计算某人该月得出勤天数。 只要刷卡一次，就计为出勤一次。
        public static int getActualAttendanceDays(string _job_Number,string _year_and_month) {
            string procedureName = "get_Actual_AR_Days";
            int v_Actual_AR_Days = 0;
            //分析刚刚导入的数据。
            OracleParameter parma_JN = new OracleParameter("v_job_number", OracleDbType.Varchar2, ParameterDirection.Input);
            OracleParameter param_year_and_Month = new OracleParameter("v_year_and_month", OracleDbType.Varchar2, ParameterDirection.Input);
            OracleParameter param_AR_Days = new OracleParameter("v_AR_Days", OracleDbType.Int32, ParameterDirection.Output);
            parma_JN.Value = _job_Number;
            param_year_and_Month.Value = _year_and_month;
            OracleParameter[] parameters = new OracleParameter[3] { parma_JN, param_year_and_Month, param_AR_Days };
            OracleHelper oH = Tools.OracleHelper.getBaseDao();
            int j = oH.ExecuteNonQuery(procedureName, parameters);
            v_Actual_AR_Days = Int32.Parse(parameters[2].Value.ToString());
            return v_Actual_AR_Days;
        }
        #endregion
    }
}
