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
            string result = "unique";
            string sqlStr = String.Format(@"
                                        SELECT 1 
                                        FROM Employees emps
                                        WHERE Name = '{0}'
                                    ",name);
            DataTable dt = Tools.OracleDaoHelper.getDTBySql(sqlStr);
            switch (dt.Rows.Count) {
                case 0:
                    return "用户不存在！";
                case 1:
                    return result;
                default:
                    return "系统中存在同名用户: " + dt.Rows.Count + "个";
            }
        }
    }
}
