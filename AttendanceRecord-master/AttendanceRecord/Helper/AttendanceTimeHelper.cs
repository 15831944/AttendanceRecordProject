using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
namespace AttendanceRecord.Helper
{
    /// <summary>
    /// 该类用于解决考勤的
    /// </summary>
    public class AttendanceTimeHelper
    {
        /// <summary>
        /// 判断是否迟到
        /// </summary>
        /// <param name="timeStr"></param>
        /// <returns></returns>
        public static bool isComeLate(string timeStr) {
            int hour = int.Parse(timeStr.Substring(0, 2));
            if (hour <= 7) return false;
            if (8 == hour) {
                int minute = int.Parse(timeStr.Substring(3, 2));
                if (0 == minute) return false;
                return true;
            }
            //9点以后，为迟到。
            return true;
        }
        /// <summary>
        /// 17:00下班
        /// </summary>
        /// <param name="timeStr"></param>
        /// <returns></returns>
        public static bool isLeaveEarly(string timeStr) {
            int hour = int.Parse(timeStr.Substring(0, 2));
            if (hour >= 17 ) return false;
            return true;
        }
        public static void judgeIfComeLateOrLeaveEarly(string timeStr, out bool comeLate  , out bool leaveEarly) {
            //先获取fpt_first_time,fpt_last_time;
            string fpt_first_time;
            string fpt_last_time;
            AttendanceRHelper.GetFPTime(timeStr, out fpt_first_time, out fpt_last_time);
            if (string.IsNullOrEmpty(fpt_last_time) && string.IsNullOrEmpty(fpt_last_time)) {
                //无所谓  迟到，早退。
                comeLate = false;
                leaveEarly = false;
                return;
            }
            if (string.IsNullOrEmpty(fpt_last_time)) {
                //没刷下班卡
                leaveEarly = false;
                //判断早上是否迟到
                comeLate = isComeLate(fpt_first_time);
                return;
            }
            if (string.IsNullOrEmpty(fpt_first_time)) {
                comeLate = false;
                leaveEarly = isLeaveEarly(fpt_last_time);
                return;
            }
            comeLate = isComeLate(fpt_first_time);
            leaveEarly = isLeaveEarly(fpt_last_time);
        }
    }
}
