using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Oracle.DataAccess.Client;
using System.Data;
using Tools;
using System.Windows.Forms;

namespace AttendanceRecord.View
{
    /// <summary>
    /// 此类用来统计延时时间。
    /// </summary>
    public class V_Delay_Time
    {
        #region 注释部分
        /*
        /// <summary>
        /// 具体日期
        /// </summary>
        private string _year_month_day;
        private string job_number;

        public string Job_number
        {
            get
            {
                return job_number;
            }

            set
            {
                job_number = value;
            }
        }

        public string Year_month_day
        {
            get
            {
                return _year_month_day;
            }

            set
            {
                _year_month_day = value;
            }
        }

        public V_Delay_Time(string year_month_day, string job_number)
        {
            this._year_month_day = year_month_day;
            this.Job_number = job_number;
        }
        */
        #endregion
        #region 计算某员工，指定月份的延时时间。
        public static double getDelayTotalTime(string _jN,string _year_and_month) {
            string procedureName = "GET_Delay_Total_Time";
            OracleParameter param_JN = new OracleParameter("v_JOB_NUMBER", OracleDbType.Varchar2, ParameterDirection.Input);
            OracleParameter param_year_and_month = new OracleParameter("v_year_and_month", OracleDbType.Varchar2, ParameterDirection.Input);
            OracleParameter param_vacation_hours = new OracleParameter("v_Delay_Total_Time", OracleDbType.Double, ParameterDirection.Output);
            OracleParameter[] parameters = new OracleParameter[3] { param_JN, param_year_and_month, param_vacation_hours };
            parameters[0].Value = _jN;
            parameters[1].Value = _year_and_month;
            OracleHelper oH = OracleHelper.getBaseDao();
            oH.ExecuteNonQuery(procedureName, parameters);
            return double.Parse(parameters[2].Value.ToString());
        }
        #endregion

    }
}
