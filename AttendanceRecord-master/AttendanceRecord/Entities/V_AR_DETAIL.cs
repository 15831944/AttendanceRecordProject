using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using Tools;
using Oracle.DataAccess.Client;
namespace AttendanceRecord.Entities
{
    public class V_AR_DETAIL
    {
       private string  _start_date;
        private string _end_date;
        private string _tabulation_time;
        private string _fingerprint_date;
        private string _job_number;
        private string _name;
        private string _dept;
        private string _fpt_first_time;
        private string _fpt_last_time;
        private string _come_Late_Num;
        private string _leave_early_num;

        public string Start_date
        {
            get
            {
                return _start_date;
            }

            set
            {
                _start_date = value;
            }
        }

        public string End_date
        {
            get
            {
                return _end_date;
            }

            set
            {
                _end_date = value;
            }
        }

        public string Tabulation_time
        {
            get
            {
                return _tabulation_time;
            }

            set
            {
                _tabulation_time = value;
            }
        }

        public string Fingerprint_date
        {
            get
            {
                return _fingerprint_date;
            }

            set
            {
                _fingerprint_date = value;
            }
        }

        public string Job_number
        {
            get
            {
                return _job_number;
            }

            set
            {
                _job_number = value;
            }
        }
        /// <summary>
        /// 获取考勤表中原始的工号。
        /// </summary>
        public string _Previous_Job_number
        {
            get
            {
                if (_job_number.Length >= 12)
                {
                    return _job_number.Substring(9);
                }
                return _job_number;
            }
        }
        public string Name
        {
            get
            {
                return _name;
            }

            set
            {
                _name = value;
            }
        }

        public string Dept
        {
            get
            {
                return _dept;
            }

            set
            {
                _dept = value;
            }
        }

        public string Fpt_first_time
        {
            get
            {
                return _fpt_first_time;
            }

            set
            {
                _fpt_first_time = value;
            }
        }

        public string Fpt_last_time
        {
            get
            {
                return _fpt_last_time;
            }

            set
            {
                _fpt_last_time = value;
            }
        }

        public string Come_Late_Num
        {
            get
            {
                return _come_Late_Num;
            }

            set
            {
                _come_Late_Num = value;
            }
        }

        public string Leave_early_num
        {
            get
            {
                return _leave_early_num;
            }

            set
            {
                _leave_early_num = value;
            }
        }
        #region 获取单个用户的月刷卡详细信息。
        public static List<V_AR_DETAIL> get_V_AR_Detail_List(string YearAndMonthStr){
            string sqlStr = string.Format(@"select 
                                              to_char(start_date,'yyyy-MM-dd') as start_date, 
                                              to_char(end_date,'yyyy-MM-dd') as end_date, 
                                              to_char(tabulation_time,'yyyy-MM-dd') as tabulation_time , 
                                              to_char(fingerprint_date,'yyyy-MM-dd') as finger_print_date, 
                                              cast(job_number as varchar(16)) as job_number, 
                                              name, 
                                              dept, 
                                              TO_CHAR(fpt_first_time,'yyyy-MM-dd') as fpt_first_time, 
                                              to_char(fpt_last_time,'yyyy-MM-dd') as fpt_last_time,
                                              cast (come_late_num as varchar2(10)) as come_late_num,
                                              cast(leave_early_num as varchar2(10)) as leave_early_num
                                            from v_AR_Detail v_AR_D
                                            where TRUNC(v_AR_D.finger_print_date,'MM') =TO_DATE( '{0}','yyyy-MM')
                                            order by v_AR_D.job_Number asc,
                                               v_AR_D.finger_print_date asc",
                                               YearAndMonthStr);
            return ConvertHelper<V_AR_DETAIL>.ConvertToList(OracleDaoHelper.getDTBySql(sqlStr));
        }
        #endregion
        #region 获取单个用户的月刷卡详细信息。
        public static V_AR_DETAIL get_V_AR_Detail(string YearAndMonthStr)
        {
            string sqlStr = string.Format(@"
                                                SELECT TEMP.start_date,
                                                        TEMP.end_date,
                                                        TEMP.tabulation_time,
                                                        TEMP.finger_print_date,
                                                        TEMP.job_number,
                                                        TEMP.name,
                                                        TEMP.dept,
                                                        TEMP.fpt_first_time,
                                                        TEMP.fpt_last_time,
                                                        TEMP.come_late_num,
                                                        TEMP.leave_early_num
                                                FROM (  
                                                        select 
                                                              rownum AS row_num,
                                                              to_char(start_date,'yyyy-MM-dd') as start_date, 
                                                              to_char(end_date,'yyyy-MM-dd') as end_date, 
                                                              to_char(tabulation_time,'yyyy-MM-dd') as tabulation_time , 
                                                              to_char(finger_print_date,'yyyy-MM-dd') as finger_print_date, 
                                                              cast(job_number as varchar(16)) as job_number, 
                                                              name, 
                                                              dept, 
                                                              TO_CHAR(fpt_first_time,'yyyy-MM-dd') as fpt_first_time, 
                                                              to_char(fpt_last_time,'yyyy-MM-dd') as fpt_last_time,
                                                              cast (come_late_num as varchar2(10)) as come_late_num,
                                                              cast(leave_early_num as varchar2(10)) as leave_early_num
                                                            from v_AR_Detail v_AR_D
                                                            where TRUNC(v_AR_D.finger_print_date,'MM') =TO_DATE( '{0}','yyyy-MM')
                                                            order by v_AR_D.job_Number asc,
                                                               v_AR_D.finger_print_date asc
                                            ) TEMP 
                                            WHERE TEMP.row_num = 1",
                                               YearAndMonthStr);
            System.Data.DataTable dt = OracleDaoHelper.getDTBySql(sqlStr);
            if (dt.Rows.Count == 0) return null;
            return ConvertHelper<V_AR_DETAIL>.ConvertToList(dt)[0];
        }
        #endregion
        #region 获取单个用户的月刷卡详细信息。
        public static List<V_AR_DETAIL> get_V_AR_Detail_List_By_YAndM_And_PreJN(string YearAndMonthStr,string prefix_JN)
        {
            string sqlStr = string.Format(@"select 
                                              to_char(start_date,'yyyy-MM-dd') as start_date, 
                                              to_char(end_date,'yyyy-MM-dd') as end_date, 
                                              to_char(tabulation_time,'yyyy-MM-dd') as tabulation_time , 
                                              to_char(fingerprint_date,'yyyy-MM-dd') as fingerprint_date, 
                                              cast(job_number as varchar(16)) as job_number, 
                                              name, 
                                              dept, 
                                              TO_CHAR(fpt_first_time,'yyyy-MM-dd') as fpt_first_time, 
                                              to_char(fpt_last_time,'yyyy-MM-dd') as fpt_last_time,
                                              cast (come_late_num as varchar2(10)) as come_late_num,
                                              cast(leave_early_num as varchar2(10)) as leave_early_num
                                            from v_AR_Detail v_AR_D
                                            where TRUNC(v_AR_D.fingerprint_date,'MM') =TO_DATE( '{0}','yyyy-MM')
                                            and substr(v_AR_D.job_number,1,9) = '{1}'
                                            order by v_AR_D.job_Number asc,
                                               v_AR_D.fingerprint_date asc",
                                               YearAndMonthStr,prefix_JN);
            return ConvertHelper<V_AR_DETAIL>.ConvertToList(OracleDaoHelper.getDTBySql(sqlStr));
        }
        #endregion
        #region
        public static List<int> getMonthAL(string YearAndMonthStr)
        {
            string sqlStr = String.Format(@"select DISTINCT(SUBSTR(TO_CHAR(v_aR_D.Fingerprint_Date,'YYYY-MM-DD'),9,2)) AS  Day
                                           from v_ar_detail v_aR_D
                                           WHERE TRUNC(v_aR_D.FINGERPRINT_DATE,'MM') = To_DATE('{0}','yyyy-MM')
                                            ORDER BY day",
                                            YearAndMonthStr
                                           );
            DataTable dt = OracleDaoHelper.getDTBySql(sqlStr);
            List<int> dayList = new List<int>();
            for (int j=0;j<= dt.Rows.Count-1;j++) {
                dayList.Add(int.Parse(dt.Rows[j]["Day"].ToString()));
            }
            return dayList;
        }
        #endregion
        #region
        public static List<int> getMonthAL_By_YAndM_And_Prefix_JN(string YearAndMonthStr,string prefix_Job_Number)
        {
            string sqlStr = String.Format(@"select DISTINCT(SUBSTR(TO_CHAR(v_aR_D.finger_print_date,'YYYY-MM-DD'),9,2)) AS  Day
                                           from v_ar_detail v_aR_D
                                           WHERE TRUNC(v_aR_D.finger_print_date,'MM') = To_DATE('{0}','yyyy-MM')
                                            and substr(v_aR_D.job_number,1,9) = '{1}'
                                            ORDER BY day",
                                            YearAndMonthStr,prefix_Job_Number
                                           );
            DataTable dt = OracleDaoHelper.getDTBySql(sqlStr);
            List<int> dayList = new List<int>();
            for (int j = 0; j <= dt.Rows.Count - 1; j++)
            {
                dayList.Add(int.Parse(dt.Rows[j]["Day"].ToString()));
            }
            return dayList;
        }
        #endregion
        #region
        /// <summary>
        /// 获取指定月份的 出勤日。
        /// </summary>
        /// <param name="YearAndMonthStr"></param>
        /// <returns></returns>
        public static List<int> getARDaysOfSpecificMonth(string YearAndMonthStr)
        {
            string sqlStr = String.Format(@"select distinct(to_number(to_char(fingerprint_date,'dd'))) AS Day
                                               from attendance_record
                                               where trunc(fingerprint_date,'MM') = to_date('{0}','yyyy-MM')
                                               order by day",
                                            YearAndMonthStr
                                           );
            DataTable dt = OracleDaoHelper.getDTBySql(sqlStr);
            List<int> dayList = new List<int>();
            for (int j = 0; j <= dt.Rows.Count - 1; j++)
            {
                dayList.Add(int.Parse(dt.Rows[j]["Day"].ToString()));
            }
            return dayList;
        }
        #endregion
        #region
        /*
        public static List<int> getMonthAL_By_YAndM(string YearAndMonthStr)
        {
            string sqlStr = String.Format(@"select DISTINCT(SUBSTR(TO_CHAR(v_aR_D.Fingerprint_Date,'YYYY-MM-DD'),9,2)) AS  Day
                                           from v_ar_detail v_aR_D
                                           WHERE TRUNC(v_aR_D.FINGERPRINT_DATE,'MM') = To_DATE('{0}','yyyy-MM')
                                            ORDER BY day",
                                            YearAndMonthStr
                                           );
            DataTable dt = OracleDaoHelper.getDTBySql(sqlStr);
            List<int> dayList = new List<int>();
            for (int j = 0; j <= dt.Rows.Count - 1; j++)
            {
                dayList.Add(int.Parse(dt.Rows[j]["Day"].ToString()));
            }
            return dayList;
        }
        */
        public static List<int> getDayListOfTheSpecificMonth(string YearAndMonthStr)
        {
            string sqlStr = String.Format(@"select distinct(to_char(finger_print_date,'DD')) day_str
                                              from Attendance_Record_Final
                                              where trunc(finger_print_date,'MM') = to_date('{0}','yyyy-MM')
                                              order by day_str asc",
                                            YearAndMonthStr
                                           );
            DataTable dt = OracleDaoHelper.getDTBySql(sqlStr);
            List<int> dayList = new List<int>();
            for (int j = 0; j <= dt.Rows.Count - 1; j++)
            {
                dayList.Add(int.Parse(dt.Rows[j]["day_str"].ToString()));
            }
            return dayList;
        }
        #endregion
        #region
        public static List<int> getDaysOfARSummaryByYearAndMonth(string YearAndMonthStr)
        {
            string sqlStr = String.Format(@"select DISTINCT(to_char(fingerprint_date,'dd')) AS  Day
                                           from V_A_R_Summary_DETAIL v_A_R_S_D
                                           WHERE TRUNC(v_A_R_S_D.FINGERPRINT_DATE,'MM') = To_DATE('{0}','yyyy-MM')
                                            ORDER BY day",
                                            YearAndMonthStr
                                           );
            DataTable dt = OracleDaoHelper.getDTBySql(sqlStr);
            List<int> dayList = new List<int>();
            for (int j = 0; j <= dt.Rows.Count - 1; j++)
            {
                dayList.Add(int.Parse(dt.Rows[j]["Day"].ToString()));
            }
            return dayList;
        }
        #endregion

        public static List<V_AR_DETAIL> get_V_AR_Detail_By_Specific_Day(string date) 
        {
            string proceName = "PKG_AR_Detail.get_staffs_base_info"; 
            OracleParameter param_date_str = new OracleParameter("v_year_and_month_str", OracleDbType.Varchar2, ParameterDirection.Input);
            OracleParameter param_cur_result = new OracleParameter("v_cur_result", OracleDbType.RefCursor, ParameterDirection.ReturnValue);
            param_date_str.Value = date;
            param_date_str.Size = 20;
            OracleParameter[] parameters = new OracleParameter[2] { param_cur_result, param_date_str  };
            OracleHelper oH = OracleHelper.getBaseDao();
            DataTable dt = oH.getDT(proceName, parameters);
            return ConvertHelper<V_AR_DETAIL>.ConvertToList(dt);
        }
        /*
        /// <summary>
        /// 获取 某月某考勤机的  部门，工号，姓名。
        /// </summary>
        /// <param name="attendance_machine_flag"></param>
        /// <param name="date"></param>
        /// <returns></returns>
        public static List<V_AR_DETAIL> get_V_AR_Detail_By_Attendance_Machine_Flag_And_Specific_Day(string attendance_machine_flag,string date)
        {
            /*string proceName = "PKG_AR_Detail.get_Staffs_BI_by_AMFlag_YMStr";
            OracleParameter param__attendance_machine_flag = new OracleParameter("v_attendance_machine_flag", OracleDbType.Varchar2, ParameterDirection.Input);
            OracleParameter param_date_str = new OracleParameter("v_year_and_month_str", OracleDbType.Varchar2, ParameterDirection.Input);
            OracleParameter param_cur_result = new OracleParameter("v_cur_result", OracleDbType.RefCursor, ParameterDirection.ReturnValue);
            param__attendance_machine_flag.Value = attendance_machine_flag;
            param_date_str.Value = date;
            param_date_str.Size = 20;
            OracleParameter[] parameters = new OracleParameter[3] { param_cur_result, param__attendance_machine_flag, param_date_str };
            OracleHelper oH = OracleHelper.getBaseDao();
            DataTable dt = oH.getDT(proceName, parameters);
            */
            /*
            string sqlStr = string.Format(@"select distinct dept,job_number,name
                                                from Attendance_Record
                                                where substr(job_number,1,1) in ({0})
                                                and trunc(fingerprint_date,'MM') = to_date('{1}','yyyy-MM')
                                                order by job_number asc",
                                                attendance_machine_flag,
                                                date);
            System.Data.DataTable dt = OracleDaoHelper.getDTBySql(sqlStr);
            return ConvertHelper<V_AR_DETAIL>.ConvertToList(dt);
        }
        */
        /// <summary>
        /// 获取 某月某考勤机的  部门，工号，姓名。
        /// </summary>
        /// <param name="attendance_machine_flag"></param>
        /// <param name="date"></param>
        /// <returns></returns>
        public static List<V_AR_DETAIL> get_V_AR_BaseInfo_By_Attendance_Machine_Flag_And_Specific_Month(string attendance_machine_flag, string year_month_str)
        {
            string sqlStr = string.Format(@"select 
                                                distinct dept,
                                                    job_number,
                                                    name
                                                from Attendance_Record_Final
                                                where substr(job_number,1,1) in ({0})
                                                and trunc(finger_print_date,'MM') = to_date('{1}','yyyy-MM')
                                                order by job_number asc",
                                                attendance_machine_flag,
                                                year_month_str);
            System.Data.DataTable dt = OracleDaoHelper.getDTBySql(sqlStr);
            return ConvertHelper<V_AR_DETAIL>.ConvertToList(dt);
        }
        /// <summary>
        /// 获取员工的基本信息。
        /// </summary>
        /// <param name="date"></param>
        /// <returns></returns>
        public static List<V_AR_DETAIL> get_V_A_R_Summary_Base_Info_By_Specific_Day(string dateStr)
        {
            string sqlStr = string.Format(@"select distinct 
                                                          dept,
                                                          job_number,
                                                          name 
                                          from Attendance_Record_Summary
                                            where fingerprint_date = to_date('{0}','yyyy-MM-dd')
                                          order by job_number asc",dateStr);
            System.Data.DataTable dt = OracleDaoHelper.getDTBySql(sqlStr);
            return ConvertHelper<V_AR_DETAIL>.ConvertToList(dt);          
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="date"></param>
        /// <param name="prefix_Job_Number"></param>
        /// <returns></returns>
        public static List<V_AR_DETAIL> get_V_AR_Detail_By_Specific_Day_And_PreJN(string date,string prefix_Job_Number)
        {
            OracleParameter param_date_str = new OracleParameter("v_date_str", OracleDbType.Varchar2, ParameterDirection.Input);
            OracleParameter param_prefix_job_number = new OracleParameter("v_prefix_job_number", OracleDbType.Varchar2, ParameterDirection.Input);
            OracleParameter param_cur_result = new OracleParameter("v_cur_result", OracleDbType.RefCursor, ParameterDirection.Output);
            param_date_str.Value = date;
            param_prefix_job_number.Value = prefix_Job_Number;
            param_date_str.Size = 20;
            param_prefix_job_number.Size = 20;
            OracleParameter[] parameters = new OracleParameter[3] {param_date_str,param_prefix_job_number,param_cur_result };
            OracleHelper oH = OracleHelper.getBaseDao();
            string proceName = "PKG_AR_Detail.get_ar_detail";
            DataTable dt = oH.getDT(proceName, parameters);
            return ConvertHelper<V_AR_DETAIL>.ConvertToList(dt);
        }
    }
}
