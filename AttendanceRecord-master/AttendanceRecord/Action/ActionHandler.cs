using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
namespace AttendanceRecord.Action
{
    public class ActionHandler
    {
        /// <summary>
        /// 获取Action
        /// </summary>
        /// <returns></returns>
        public static bool getAction(string _require_Action) {
            if ("admin".Equals(Program._userInfo.Action,StringComparison.CurrentCultureIgnoreCase)) {
                return true;
            }
            switch (_require_Action) {
                case @"Common":
                    return true;
                default:
                    if (!_require_Action.Equals(Program._userInfo.Action, StringComparison.CurrentCultureIgnoreCase)) {
                        MessageBox.Show("权限不足","提示:", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return false;
                    }
                    return true;
            }
        }
    }
}
