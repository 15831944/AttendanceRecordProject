using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using Tools;
namespace AttendanceRecord.Entities
{
    public class CheckUsers
    {
        public static bool ifExistsUserName(string name) {
            string sqlStr = string.Format(@"SELECT 1 FROM USER_INFO WHERE User_Name = '{0}'", name);
            return OracleDaoHelper.getDTBySql(sqlStr).Rows.Count > 0 ? true : false;
        }
        public static bool isPasswordRight(string userName, string password) {
            string sqlStr = String.Format(@"SELECT 1 FROM USER_INFO WHERE User_Name = '{0}'
                                                AND Password = '{1}'", userName, password);
            int rows_num = 0;
            rows_num = OracleDaoHelper.getDTBySql(sqlStr).Rows.Count;
            return rows_num > 0 ? true : false;
        }
        public static User_Info getUserInfo(string userName) {
            string sqlStr = String.Format(@"SELECT user_name,
                                                    password,
                                                    TO_CHAR(update_time,'yyyy-MM-dd') AS UPDATE_TIME,
                                                    department,
                                                     ACTION
                                                FROM USER_INFO WHERE User_Name = '{0}'", userName
                                                );
            List<User_Info> userInfoList = ConvertHelper<User_Info>.ConvertToList(OracleDaoHelper.getDTBySql(sqlStr));
            return userInfoList[0];
        }
    }
}
