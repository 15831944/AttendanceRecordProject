using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using Tools;
namespace AttendanceRecord.View
{
    public class V_Staff
    {
        private string _year_and_month;

        private string _name;

        private string _job_number;

        public string Year_and_month
        {
            get
            {
                return _year_and_month;
            }

            set
            {
                _year_and_month = value;
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

        public V_Staff(string _year_and_month, string _name)
        {
            this.Year_and_month = _year_and_month;
            this.Name = _name;
        }

        public V_Staff() {

        }
        /// <summary>
        /// 返回该姓名的员工的姓名，工号，部门，最早上班时间，最晚下班时间，总出勤天数。
        /// </summary>
        /// <returns></returns>
        public DataTable getARInfo() {
            string sqlStr = string.Format(@"
                                          select name,
                                                 job_number,
                                                 dept,
                                                 min(fpt_first_time) min_fpt_first_time,
                                                 max(fpt_last_time) max_fpt_last_time,
                                                 sum(case when (fpt_first_time is null and fpt_last_time is null) then 0 else 1 end) as Actual_AR_Days
                                          from Attendance_Record AR
                                          where name = '{0}' 
                                          and trunc(AR.Fingerprint_Date,'MM') = to_date('{1}','yyyy-MM')
                                          group by name,job_number,dept",
                                            this.Name,
                                            this.Year_and_month);
            return OracleDaoHelper.getDTBySql(sqlStr);
        }

    }
}
