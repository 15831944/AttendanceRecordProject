using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using Tools;
using AttendanceRecord.View;

namespace AttendanceRecord.Entities
{
    /// <summary>
    /// 用于补卡的视图
    /// </summary>
    public class V_FillUp
    {
        private string _name;
        private string _day;
        private string _time;
        public V_FillUp(string name, string day, string time) { this._name = name; this._day = day; this._time = time; }

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

        public string Time
        {
            get
            {
                return _time;
            }

            set
            {
                _time = value;
            }
        }

        public int getFillUpRecordTimes() {
            string sqlStr = string.Format(@"select wm_concat(to_char(fill_up_remark))
                                              from attendance_record
                                              where name = '{0}'
                                              and trunc(fingerprint_date,'MM') = to_date('{1}','yyyy-MM')
                                              and fill_up_remark is not null",
                                              this._name,
                                              this._day.Substring(0, 7));
            DataTable dt = OracleDaoHelper.getDTBySql(sqlStr);
            if (dt.Rows.Count == 0) return 0;
            return StringHelper.getSpecificChar(dt.Rows[0][0].ToString(), ";");
        }

        public DataTable getARRecordToFillUpByName() {
            string sqlStr = string.Format(@"SELECT JOB_NUMBER  AS ""工号"",
                                                 NAME AS ""姓名"",
                                                 DEPT  AS ""部门"",
                                                 FPT_FIRST_TIME AS ""上班时间"",
                                                 FPT_LAST_TIME  AS ""下班时间"",
                                                 FILL_UP_REMARK AS ""备注""
                                          FROM Attendance_Record
                                          WHERE name = '{0}'
                                          AND TRUNC(FINGERPRINT_DATE,'MM')= TO_DATE('{1}','yyyy-MM')
                                          ORDER BY FINGERPRINT_DATE DESC 
                                          ", this._name,
                                          this._day.Substring(0, 7));
            DataTable dt = OracleDaoHelper.getDTBySql(sqlStr);
            return dt;
        }

        #region 判断早上是否有刷卡记录
        public bool ifNotHaveRecordOfMorning() {
            string sqlStr = string.Format(@"SELECT 1 FROM Attendance_Record 
                                                    WHERE NAME = '{0}'  
                                                    AND TRUNC(FINGERPRINT_DATE,'DD')= TO_DATE('{1}','yyyy-MM-dd')
                                                     AND FPT_FIRST_TIME IS NULL",
                                                    this.Name,
                                                    this.Day);
            DataTable dt = OracleDaoHelper.getDTBySql(sqlStr);
            return dt.Rows.Count > 0 ? true : false;
        }
        #endregion

        #region 判断下午是否有刷卡记录
        public bool ifNotHaveRecordOfAfternoon() {
            string sqlStr = string.Format(@"SELECT 1 FROM Attendance_Record
                                                WHERE NAME = '{0}'
                                                    AND TRUNC(FINGERPRINT_DATE,'DD')= TO_DATE('{1}','yyyy-MM-dd')
                                                    AND FPT_LAST_TIME IS NULL",
                                                    this.Name,
                                                    this.Day
                                                        );
            DataTable dt = OracleDaoHelper.getDTBySql(sqlStr);
            return dt.Rows.Count > 0 ? true : false;
        }
        #endregion

        #region 更新某时间段的刷卡记录.
        public bool updateTheRecord() {
            string sqlStr_update_FPT_First_Time = string.Format(@"UPDATE Attendance_Record
                                                SET FPT_FIRST_TIME = to_date('{2}','yyyy-MM-dd HH24:MI:SS'),
                                                        FILL_UP_REMARK = FILL_UP_REMARK || '   '|| '{3}' || '   '|| to_char(sysdate,'yyyy-MM-dd HH24:MI:SS') ||';'
                                                 WHERE NAME = '{0}'
                                                    AND TRUNC(FINGERPRINT_DATE,'DD')= TO_DATE('{1}','yyyy-MM-dd')",
                                                   this._name,
                                                   this._day,
                                                   this._day + " " + this._time,
                                                   "已补上班卡");
            string sqlStr_update_FPT_Last_Time = string.Format(@"UPDATE Attendance_Record
                                                SET FPT_LAST_TIME = to_date('{2}','yyyy-MM-dd HH24:MI:SS'),
                                                    FILL_UP_REMARK = FILL_UP_REMARK || '   '|| '{3}'|| '   '|| to_char(sysdate,'yyyy-MM-dd HH24:MI:SS') ||';'
                                                 WHERE NAME = '{0}'
                                                    AND TRUNC(FINGERPRINT_DATE,'DD')= TO_DATE('{1}','yyyy-MM-dd')",
                                                  this._name,
                                                  this._day,
                                                  this._day + " " + this._time,
                                                  "已补下班卡");
            if (!ifNotHaveRecordOfAfternoon() && !ifNotHaveRecordOfMorning()) {
                System.Windows.Forms.MessageBox.Show(this._day + ": 有出勤记录，无需补卡！", "提示：", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                return false;
            }
            if (ifNotHaveRecordOfMorning() && !ifNotHaveRecordOfAfternoon()) {
                if (!ifTheTimeEarlierThanLastTime()) {
                    System.Windows.Forms.MessageBox.Show("所补上班时间: " + this._day + " " + this._time + " 必须 <= " +new V_AR(this._name,this._day).getLastTime(), "提示：", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                    return false;
                }
                OracleDaoHelper.executeSQL(sqlStr_update_FPT_First_Time);
                return true;
            }

            if (!ifNotHaveRecordOfMorning() && ifNotHaveRecordOfAfternoon()) {
                if (!ifTheTimeLaterThanFirstTime())
                {
                    System.Windows.Forms.MessageBox.Show("所补下班时间：" + this._day + " " + this._time + " 必须 >= " + new V_AR(this._name, this._day).getFirstTime(), "提示：", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                    return false;
                }
                OracleDaoHelper.executeSQL(sqlStr_update_FPT_Last_Time);
                return true;
            }
            if (ifNotHaveRecordOfMorning() && ifNotHaveRecordOfAfternoon()) {
                if (ifMorning())
                {
                    OracleDaoHelper.executeSQL(sqlStr_update_FPT_First_Time);
                    return true;
                }
                OracleDaoHelper.executeSQL(sqlStr_update_FPT_Last_Time);
                return true;
            }
            return false;
        }
        #endregion

        #region 判断时间点是早上,还是下午.
        public bool ifMorning() {
            string[] timeArray = _time.Split(':');
            int hour = int.Parse(timeArray[0]);
            int minute = int.Parse(timeArray[1]);
            int second = int.Parse(timeArray[2]);

            string[] dayArray = _day.Split('-');
            int year = int.Parse(dayArray[0]);
            int month = int.Parse(dayArray[1]);
            int day = int.Parse(dayArray[2]);

            DateTime dt = new DateTime(year, month, day, hour, minute, second);
            DateTime dtNoon = new DateTime(year, month, day, 12, 0, 0);

            if (dt < dtNoon) return true;
            return false;
        }
        #endregion
        #region 判断该时间点是否早于下班点
        public bool ifTheTimeEarlierThanLastTime(){
            string sqlStr = string.Format(@"select 1 
                                                from Attendance_Record 
                                                where name = '{0}'
                                                and fingerprint_date = to_date('{1}','yyyy-MM-dd')
                                                and  to_date('{1} {2}','yyyy-MM-dd HH24:MI:SS') < FPT_LAST_TIME",
                                                this._name,
                                                this._day,
                                                this._time);
            DataTable dt = OracleDaoHelper.getDTBySql(sqlStr);
            return dt.Rows.Count > 0 ? true : false;
        }
        #endregion
        #region 判断该时间点是否晚于上班点
        public bool ifTheTimeLaterThanFirstTime()
        {
            string sqlStr = string.Format(@"select 1 
                                                from Attendance_Record 
                                                where name = '{0}'
                                                and fingerprint_date = to_date('{1}','yyyy-MM-dd')
                                                and  to_date('{1} {2}','yyyy-MM-dd HH24:MI:SS') > FPT_FIRST_TIME",
                                                this._name,
                                                this._day,
                                                this._time);
            DataTable dt = OracleDaoHelper.getDTBySql(sqlStr);
            return dt.Rows.Count > 0 ? true : false;
        }
        #endregion
    }
}
