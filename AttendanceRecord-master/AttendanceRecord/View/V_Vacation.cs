using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using Tools;
using AttendanceRecord.Helper;
using Oracle.DataAccess.Client;
namespace AttendanceRecord.View
{
/// <summary>
/// 假期视图,主要用来计算当天,请假的时间.
/// </summary>
    public class V_Vacation
    {
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

        public V_Vacation(string job_number,string year_month_day )
        {
            this._year_month_day = year_month_day;
            this.Job_number = job_number;
        }
        #region 获取该天为事假时的起始时间点
        public string getStartTimeStr()
        {
            string sqlStr = String.Format(@"
                                            SELECT to_char(leave_start_time,'hh24:mi') AS startTimeStr
                                            FROM ASK_FOR_LEAVE a_f_l
                                            WHERE a_f_l.job_number = '{0}'
                                            and trunc(leave_start_time,'DD') = to_date('{1}','yyyy-MM-dd')
                                            ",
                                            this.Job_number,
                                            this._year_month_day);
            DataTable dt = Tools.OracleDaoHelper.getDTBySql(sqlStr);
            return dt.Rows[0]["startTimeStr"].ToString();
        }
        #endregion
        #region 获取该天为事假时的结束时间点
       public   string getEndTimeStr()
        {
            string sqlStr = String.Format(@"
                                            SELECT to_char(leave_end_time,'hh24:mi') AS endTimeStr
                                            FROM ASK_FOR_LEAVE a_f_l
                                            WHERE a_f_l.job_number = '{0}'
                                            and trunc(leave_end_time,'DD') = to_date('{1}','yyyy-MM-dd')
                                            ",
                                            this.Job_number,
                                            this._year_month_day);
            DataTable dt = Tools.OracleDaoHelper.getDTBySql(sqlStr);
            return dt.Rows[0]["endTimeStr"].ToString();
        }
        #endregion
        #region 判断当天是否请假.
        public bool ifExistsVacation() {


            bool result = false;
            //1.先判断当天是否为公司休息日.
            result = Have_A_Rest_Helper.ifDayOfRest(this._year_month_day);
            if (result) return false;
            // 2. 判断当天是否请假了.
            //这里要写员工条件
            string sqlStr = string.Format(@"SELECT 1 
                                        FROM ASK_FOR_LEAVE A_F_L
                                        WHERE A_F_L.Job_Number = '{0}'
                                        AND TRUNC(A_F_L.Leave_start_time,'DD') = TO_DATE('{1}','yyyy-MM-dd')",
                                        this.Job_number,
                                        this._year_month_day);
            result = OracleDaoHelper.getDTBySql(sqlStr).Rows.Count > 0 ? true : false;
            return result;
        }
        #endregion
        #region 计算当天请假的时间(小时为单位)
        public int getAskForLeaveDays() {
            string sqlStr = String.Format(@"select(leave_end_Time - leave_start_time) * 24 leave_Hours
                                                    from Ask_For_Leave
                                                    where Job_Number = '{0}'
                                                    and trunc(leave_start_time, 'DD') = to_date('{1}', 'yyyy-MM-dd')",
                                                    this.Job_number,
                                                    this.Year_month_day);
            return Int32.Parse(OracleDaoHelper.getDTBySql(sqlStr).Rows[0]["leave_Hours"].ToString());
        }
        #endregion
        #region 计算某个月份总共请假的时间
        public static  int getTotalTime(string _job_number,string _year_month) {
            string procedure_name = "GET_Total_TIME_For_A_F_L";
            OracleParameter param_JN = new OracleParameter("v_JOB_NUMBER", OracleDbType.Varchar2, ParameterDirection.Input);
            OracleParameter param_year_and_month = new OracleParameter("v_year_and_month", OracleDbType.Varchar2, ParameterDirection.Input);
            OracleParameter param_vacation_total_time = new OracleParameter("v_vacation_total_time", OracleDbType.Int32, ParameterDirection.Output);
            OracleParameter[] parameters = new OracleParameter[3] { param_JN,param_year_and_month, param_vacation_total_time };
            parameters[0].Value = _job_number;
            parameters[1].Value = _year_month;
            OracleHelper oH = OracleHelper.getBaseDao();
            oH.ExecuteNonQuery(procedure_name, parameters);
            return int.Parse(parameters[2].Value.ToString());
        }
        #endregion
        #region 计算某个月份总共请假的时间
        public static int getTotalTimeByName(string name, string _year_month)
        {
            string procedure_name = "GET_T_T_For_A_F_L_By_Name";
            OracleParameter param_JN = new OracleParameter("v_Name", OracleDbType.Varchar2, ParameterDirection.Input);
            OracleParameter param_year_and_month = new OracleParameter("v_year_and_month", OracleDbType.Varchar2, ParameterDirection.Input);
            OracleParameter param_vacation_total_time = new OracleParameter("v_vacation_total_time", OracleDbType.Int32, ParameterDirection.Output);
            OracleParameter[] parameters = new OracleParameter[3] { param_JN, param_year_and_month, param_vacation_total_time };
            parameters[0].Value = name;
            parameters[1].Value = _year_month;
            OracleHelper oH = OracleHelper.getBaseDao();
            oH.ExecuteNonQuery(procedure_name, parameters);
            return int.Parse(parameters[2].Value.ToString());
        }
        #endregion
    }
}
