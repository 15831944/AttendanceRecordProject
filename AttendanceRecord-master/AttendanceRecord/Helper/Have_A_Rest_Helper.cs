using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using Tools;
namespace AttendanceRecord.Helper
{
    /// <summary>
    /// 用于判断某天是否为休息日。
    /// </summary>
    public class Have_A_Rest_Helper
    {
        #region 如果上班人数小于99人，认为该日比定为休息日。
        public static  bool ifDayOfRestAutomaticAnalysis(string year_month_day) {
            string sqlStr = String.Format(@"
                                          select 1
                                            from Attendance_Record AR
                                            where  (AR.Fpt_First_Time IS NOT NULL OR AR.Fpt_Last_Time IS NOT NULL)
                                            AND trunc(AR.fingerprint_date,'DD') = to_date( '{0}','yyyy-MM-dd')
                                            having count(1) < 99",year_month_day);
            DataTable dt = OracleDaoHelper.getDTBySql(sqlStr);
            return dt.Rows.Count > 0 ? true : false;
        }
        #endregion
        /// <summary>
        /// 判断 给定的日期字符串，在Rest_day中是否为休息日。
        /// </summary>
        /// <param name="year_month_day"></param>
        /// <returns></returns>
        public static bool ifDayOfRest(string year_month_day) {
            string sqlStr = string.Format(@"SELECT 1 
                                            FROM Rest_Day RD
                                            WHERE rest_day = TO_DATE('{0}','yyyy-MM-dd')",year_month_day);
            return OracleDaoHelper.getDTBySql(sqlStr).Rows.Count > 0 ? true : false;
        }
        /// <summary>
        /// 获取指定月份：为休息日的  dd部分。
        /// </summary>
        /// <param name="year_and_month_str"></param>
        /// <returns></returns>
        public static DataTable getDaysOfRestDay(string year_and_month_str) {
            string sqlStr = string.Format(@"SELECT TO_NUMBER(TO_CHAR(rest_day,'dd')) AS dd 
                                                FROM Rest_Day RD
                                                WHERE trunc(rest_day,'MM') = to_date('{0}','yyyy-MM')",
                                                year_and_month_str);
            return OracleDaoHelper.getDTBySql(sqlStr);
        }
    }
}
