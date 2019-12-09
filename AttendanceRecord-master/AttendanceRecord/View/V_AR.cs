using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Tools;
namespace AttendanceRecord.View
{
    /// <summary>
    /// 此视图用来查询当日的上班，下班时间。
    /// </summary>
    public class V_AR
    {
        private string _name;
        private string _day;

        public V_AR(string _name, string _day)
        {
            this._name = _name;
            this._day = _day;
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

        public string Day
        {
            get
            {
                return _day;
            }

            set
            {
                _day = value;
            }
        }

        /// <summary>
        /// 获取上班时间
        /// </summary>
        /// <returns></returns>
        public string getFirstTime() {
            string sqlStr = string.Format(@"select to_char(fpt_first_time,'yyyy-MM-dd hh24:mi:ss')
                                            from Attendance_Record 
                                            where name = '{0}'
                                            and fingerprint_date = to_date('{1}','yyyy-MM-dd')",
                                            this._name,
                                            this._day);
            return OracleDaoHelper.getDTBySql(sqlStr).Rows[0][0].ToString();
        }


        /// <summary>
        /// 获取下班时间
        /// </summary>
        /// <returns></returns>
        public string getLastTime()
        {
            string sqlStr = string.Format(@"select to_char(fpt_last_time,'yyyy-MM-dd hh24:mi:ss')
                                            from Attendance_Record 
                                            where name = '{0}'
                                            and fingerprint_date = to_date('{1}','yyyy-MM-dd')",
                                            this._name,
                                            this._day);
            return OracleDaoHelper.getDTBySql(sqlStr).Rows[0][0].ToString();
        }

    }
}
