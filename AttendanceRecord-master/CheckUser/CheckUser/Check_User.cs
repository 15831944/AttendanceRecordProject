using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Tools;
using System.Windows.Forms;
namespace CheckUser
{
    public class CheckUsers
    {
        public static bool ifExistsUserName(string name)
        {
            string sqlStr = string.Format(@"SELECT 1 FROM USER_INFO WHERE User_Name = '{0}'", name);
            return OracleDaoHelper.getDTBySql(sqlStr).Rows.Count > 0 ? true : false;
        }
        public static bool isPasswordRight(string userName, string password)
        {
            string sqlStr = String.Format(@"SELECT 1 FROM USER_INFO WHERE User_Name = '{0}'
                                                AND Password = '{1}'", userName, password);
            int rows_num = 0;
            rows_num = OracleDaoHelper.getDTBySql(sqlStr).Rows.Count;
            return rows_num > 0 ? true : false;
        }
        public static User_Info getUserInfo(string userName)
        {
            string sqlStr = String.Format(@"SELECT user_name,
                                                    password,
                                                    TO_CHAR(update_time,'yyyy/MM/dd') AS UPDATE_TIME,
                                                    department,
                                                     ACTION
                                                FROM USER_INFO WHERE User_Name = '{0}'", userName
                                                );
            List<User_Info> userInfoList = ConvertHelper<User_Info>.ConvertToList(OracleDaoHelper.getDTBySql(sqlStr));
            return userInfoList[0];
        }
        /// <summary>
        /// admin 或  对应 权限  则返回真。
        /// </summary>
        /// <param name="_require_Action">需要的权限</param>
        /// <returns></returns>
        public static bool getAction(User_Info user_info,string _require_Action = null)
        {
            if (user_info.Action.Equals("admin", StringComparison.CurrentCultureIgnoreCase))
            {
                return true;
            }
            switch (_require_Action)
            {
                case @"Common":
                    return true;
                default:
                    if (!user_info.Action.Contains(_require_Action))
                    {
                        MessageBox.Show("权限不足!", "提示:", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return false;
                    }
                    return true;
            }
        }
    }
}
