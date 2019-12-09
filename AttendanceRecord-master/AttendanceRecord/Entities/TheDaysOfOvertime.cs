using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Tools;
using System.Data;
namespace AttendanceRecord.Entities
{
    /// <summary>
    /// 这是一个类。一个非常可爱的人。
    /// 它可以判断自身的日期是否为休息日。
    /// 其次亦可以设置自身为休息日。
    /// 他有自身的姓名
    /// 还有当日的日期。
    /// 删除某个休息日。
    /// </summary>
    public class TheDaysOfOvertime
    {
        private string _name;
        private string _the_day_of_overtime;

        public TheDaysOfOvertime(string _name, string _the_day_of_overtime)
        {
            this._name = _name;
            this._the_day_of_overtime = _the_day_of_overtime;
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

        public string The_days_of_overtime
        {
            get
            {
                return _the_day_of_overtime;
            }

            set
            {
                _the_day_of_overtime = value;
            }
        }

        public bool ifRestDay() {
            string sqlStr = string.Format(@"select 1 
                                            from Rest_Day
                                            where name = '{0}'
                                            and trunc(rest_day,'DD') = to_date('{1}','yyyy-MM-dd')",
                                            _name,_the_day_of_overtime);
            return OracleDaoHelper.getDTBySql(sqlStr).Rows.Count > 0 ? true : false;
        }
        /// <summary>
        /// 本月是否已经设定了加班日。
        /// </summary>
        /// <returns></returns>
        public bool ifHaveTheDayOfOvertime() {
            string sqlStr = string.Format(@"select 1 
                                            from Rest_Day
                                            where name = '{0}'
                                            and trunc(rest_day,'MM') = to_date('{1}','yyyy-MM')",
                                           _name, _the_day_of_overtime);
            return OracleDaoHelper.getDTBySql(sqlStr).Rows.Count > 0 ? true : false;

        }
        /// <summary>
        /// 
        /// </summary>
        public void addRestDay() {
            //如果不是休息日。
            if (!ifRestDay()) {
                string sqlStr = string.Format(@"INSERT INTO Rest_Day(name,rest_day,update_time)
                                                values('{0}',to_date('{1}','yyyy-MM-dd'),sysdate)",
                                                    _name,
                                                   The_days_of_overtime);
                OracleDaoHelper.executeSQL(sqlStr);
            }
        }
        /// <summary>
        /// 删除某天的休息日。
        /// </summary>
        public void delRestDay() {
            string sqlStr = string.Format(@"DELETE REST_DAY 
                                            WHERE NAME = '{0}'
                                            AND TRUNC(Rest_Day,'DD') = to_date('{1}','yyyy-MM-dd')",
                                            _name,
                                            _the_day_of_overtime);
            OracleDaoHelper.executeSQL(sqlStr);
        }
        public static DataTable getAllRestDays() {
            string sqlStr = string.Format(@"select 
                                               Name AS ""姓名"",
                                               Rest_Day AS ""休息日"",
                                               UPDATE_Time AS ""更新日期""            
                                        from Rest_Day
                                        order by  ""更新日期"" desc,""休息日"" desc");
            return OracleDaoHelper.getDTBySql(sqlStr);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public static DataTable getRestDays(string yearAndMonth)
        {
            string sqlStr = string.Format(@"select  
                                                    Name AS ""姓名"",
                                                    Rest_Day AS ""休息日"",
                                                    UPDATE_Time AS ""更新日期""   
                                              from rest_day
                                              where trunc(rest_day,'MM') = to_date('{0}','yyyy-MM')
                                              order by update_Time desc,Rest_Day desc",
                                              yearAndMonth);
            return OracleDaoHelper.getDTBySql(sqlStr);
        }
     
    }
}
