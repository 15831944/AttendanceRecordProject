using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Tools;

namespace AttendanceRecord.Entities
{
    public class V_Emp_BaseInfo
    {
        private string _job_number;
        private string _dept;
        private string _name;

        public string Job_number { get => _job_number; set => _job_number = value; }
        public string Dept { get => _dept; set => _dept = value; }
        public string Name { get => _name; set => _name = value; }

        public static List<V_Emp_BaseInfo> get_V_AR_BaseInfo_By_Attendance_Machine_Flag_And_Specific_Month(string attendance_machine_flag, string year_month_str)
        {
            string sqlStr = string.Format(@"select 
                                                distinct 
                                                    dept,
                                                    job_number,
                                                    name
                                                from Attendance_Record_Final
                                                where substr(job_number,1,1) in ({0})
                                                and trunc(finger_print_date,'MM') = to_date('{1}','yyyy-MM')
                                                order by job_number asc",
                                                attendance_machine_flag,
                                                year_month_str);
            System.Data.DataTable dt = OracleDaoHelper.getDTBySql(sqlStr);
            return ConvertHelper<V_Emp_BaseInfo>.ConvertToList(dt);
        }
    }
}
