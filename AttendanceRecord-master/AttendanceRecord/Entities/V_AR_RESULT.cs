using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using Tools;
namespace AttendanceRecord.Entities
{
    public class V_AR_RESULT
    {
        private string _fingerprint_date;
        private string _dept;
        private string _job_number;
        private string _name;
        private string _come_num;
        private string _not_finger_print;
        private string _delay_time;
        private string _come_late_num;
        private string _leave_early_num;
        private string _year_And_Month_Str;
        private string _dinner_subsidy;
        private List<V_AR_RESULT> _v_AR_Result_List;

      


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
                if (_job_number.Length==12) {
                    return _job_number.Substring(10,2);
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

        public string Come_num
        {
            get
            {
                return _come_num;
            }

            set
            {
                _come_num = value;
            }
        }

        public string Not_finger_print
        {
            get
            {
                return _not_finger_print;
            }

            set
            {
                _not_finger_print = value;
            }
        }

        public string Delay_time
        {
            get
            {
                return _delay_time;
            }

            set
            {
                _delay_time = value;
            }
        }

        public string Come_late_num
        {
            get
            {
                return _come_late_num;
            }

            set
            {
                _come_late_num = value;
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

        public string Year_And_Month_Str
        {
            get
            {
                return _year_And_Month_Str;
            }

            set
            {
                _year_And_Month_Str = value;
            }
        }

        public List<V_AR_RESULT> V_AR_Result_List
        {
            get
            {
                return _v_AR_Result_List;
            }

            set
            {
                _v_AR_Result_List = value;
            }
        }

        public string Dinner_subsidy
        {
            get
            {
                return _dinner_subsidy;
            }

            set
            {
                _dinner_subsidy = value;
            }
        }
        public V_AR_RESULT(string Year_And_Month_Str) {
            this._year_And_Month_Str = Year_And_Month_Str;
            string sqlStr = string.Format(@"
                                          select to_char(fingerprint_date,'yyyy-MM-dd') as fingerprint_date , 
                                                    dept, 
                                                    cast(job_number as varchar2(10)) as job_number, 
                                                    name, 
                                                    cast(come_num as varchar2(10)) as com_num, 
                                                    cast(not_finger_print as varchar2(10)) as not_finger_print, 
                                                    cast(delay_time as varchar2(10)) as delay_time, 
                                                    cast(come_late_num as varchar2(10)) as come_late_num, 
                                                    cast(leave_early_num as varchar2(10)) as leave_early_num,
                                                    cast(dinner_subsidy as varchar2(10)) as dinner_subsidy
                                          from v_ar_result v_ar_r
                                          where trunc(v_ar_r.fingerprint_date,'MM') = TO_DATE('{0}','yyyy-MM')
                                          order by v_ar_r.JOB_NUMBER asc", Year_And_Month_Str);
            this._v_AR_Result_List=  Tools.ConvertHelper<V_AR_RESULT>.ConvertToList(OracleDaoHelper.getDTBySql(sqlStr));
        }

        public V_AR_RESULT() { }

        #region 给出日期，获取 AR_Result
        public static List<V_AR_RESULT> get_V_AR_Result(string Year_And_Month_Str) {
            string sqlStr = string.Format(@"
                                          select to_char(fingerprint_date,'yyyy-MM-dd') as fingerprint_date , 
                                                    dept, 
                                                    cast(job_number as varchar2(10)) as job_number, 
                                                    name, 
                                                    cast(come_num as varchar2(10)) as com_num, 
                                                    cast(not_finger_print as varchar2(10)) as not_finger_print, 
                                                    cast(delay_time as varchar2(10)) as delay_time, 
                                                    cast(come_late_num as varchar2(10)) as come_late_num, 
                                                    cast(leave_early_num as varchar2(10)) as leave_early_num,
                                                    cast(dinner_subsidy as varchar2(10)) as dinner_subsidy
                                          from v_ar_result v_ar_r
                                          where trunc(v_ar_r.fingerprint_date,'MM') = TO_DATE('{0}','yyyy-MM')
                                          order by v_ar_r.JOB_NUMBER asc", Year_And_Month_Str);
            return Tools.ConvertHelper<V_AR_RESULT>.ConvertToList(OracleDaoHelper.getDTBySql(sqlStr));
        }
        #endregion
        #region 给出日期，获取 AR_Result
        /// <summary>
        /// 
        /// </summary>
        /// <param name="Year_And_Month_Str"></param>
        /// <param name="prefix_Job_Number"></param>
        /// <returns></returns>
        public static List<V_AR_RESULT> get_V_AR_Result_By_YAndM_And_PreJN(string Year_And_Month_Str,string prefix_Job_Number)
        {
            string sqlStr = string.Format(@"
                                          select to_char(fingerprint_date,'yyyy-MM-dd') as fingerprint_date , 
                                                    dept, 
                                                    cast(job_number as varchar2(10)) as job_number, 
                                                    name, 
                                                    cast(come_num as varchar2(10)) as com_num, 
                                                    cast(not_finger_print as varchar2(10)) as not_finger_print, 
                                                    cast(delay_time as varchar2(10)) as delay_time, 
                                                    cast(come_late_num as varchar2(10)) as come_late_num, 
                                                    cast(leave_early_num as varchar2(10)) as leave_early_num,
                                                    cast(dinner_subsidy as varchar2(10)) as dinner_subsidy
                                          from v_ar_result v_ar_r
                                          where trunc(v_ar_r.fingerprint_date,'MM') = TO_DATE('{0}','yyyy-MM')
                                            and substr(v_ar_r.job_number,1,9) = '{1}'
                                          order by v_ar_r.JOB_NUMBER asc", Year_And_Month_Str, prefix_Job_Number);
            return Tools.ConvertHelper<V_AR_RESULT>.ConvertToList(OracleDaoHelper.getDTBySql(sqlStr));
        }
        #endregion
        #region 获取某个员工当月的考勤汇总
        public static List<V_AR_RESULT> get_V_AR_Result_Of_Specific_JN(string YearAndMonthStr,string job_number) {
            string sqlStr = String.Format(@"select 
                                            TO_CHAR(fingerprint_date,'YYYY-MM-DD') as fingerprint_date, 
                                            dept, 
                                            CAST(job_number AS VARCHAR2(10)) AS JOB_NUMBER, 
                                            name, 
                                            cast(come_num as varchar2(10)) as come_num, 
                                            cast(not_finger_print as varchar2(10)) as not_finger_print, 
                                            cast(delay_time as varchar2(10)) as delay_time, 
                                            cast(come_late_num as varchar2(10)) as come_late_num, 
                                            cast(leave_early_num as varchar2(10)) as leave_early_num,
                                            cast(dinner_subsidy as varchar2(10)) as dinner_subsidy  
                                      from v_ar_result v_ar_r
                                      where trunc(v_ar_r.fingerprint_date,'MM') = TO_DATE('{0}','yyyy-MM')
                                      and Job_Number = '{1}'", YearAndMonthStr, job_number);
            return ConvertHelper<V_AR_RESULT>.ConvertToList(OracleDaoHelper.getDTBySql(sqlStr));
        }
        #endregion
        #region 获取某个月中员工前缀的队列
        public static Queue<int> get_Prefix_Staffs(string YearAndMonthStr)
        {
            string sqlStr = String.Format(@"select DISTINCT(SUBSTR(JOB_NUMBER,1,9)) AS Prefix_OF_Job_NUMBER
                                                from v_Ar_Result v_ar_r
                                                where trunc(v_ar_r.fingerprint_date,'MM') = TO_DATE('{0}','yyyy-MM')
                                                ORDER BY Prefix_OF_Job_NUMBER ASC",YearAndMonthStr
                                      );
            DataTable dt = OracleDaoHelper.getDTBySql(sqlStr);
            Queue<int> prefixOfJN_Queue = new Queue<int>();
            for (int i=0;i<=dt.Rows.Count-1;i++) {
                prefixOfJN_Queue.Enqueue(int.Parse(dt.Rows[i]["Prefix_OF_Job_NUMBER"].ToString()));
            }
            return prefixOfJN_Queue;
        }
        #endregion
    }
}
