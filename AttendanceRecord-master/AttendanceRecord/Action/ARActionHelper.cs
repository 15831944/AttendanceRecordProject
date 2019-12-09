using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
namespace AttendanceRecord.Action
{
    public class ARActionHelper
    {
        /// <summary>
        /// 个人权限： import;generate;   
        /// </summary>
        /// <param name="_require_Action"></param>
        /// <returns></returns>
        public static bool getAction(string _require_Action)
        {
            if (Program._userInfo.Action.Equals("admin", StringComparison.CurrentCultureIgnoreCase))
            {
                return true;
            }
            switch (_require_Action)
            {
                case @"Common":
                    return true;
                default:
                    if (!Program._userInfo.Action.Contains(_require_Action))
                    {
                        MessageBox.Show("权限不足!", "提示:", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return false;
                    }
                    return true;
            }
        }
    }
}
