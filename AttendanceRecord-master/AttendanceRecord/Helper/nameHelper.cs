using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
namespace AttendanceRecord.Helper
{
    public class nameHelper
    {
        public static string checkName(string name) {
            string sqlStr = String.Format(@"select 1
                                          from attendance_record_briefly
                                          where trunc(finger_print_date,'MM')>=trunc(add_months(sysdate,-5),'MM') 
                                          and name = '{0}'",name);
            DataTable dt = Tools.OracleDaoHelper.getDTBySql(sqlStr);
            if (dt.Rows.Count == 0) {
                return "用户不存在！";
            }
            return "用户存在";
        }
    }
}
