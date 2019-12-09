using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
namespace AttendanceRecord.Entities
{
    /// <summary>
    /// 倒入的工作安排视图类。
    /// </summary>
    public  class V_I_W_S
    {
        private string dept;
        private string date;
        //背景色。
        private Object bgColor;

        public string Dept
        {
            get
            {
                return dept;
            }

            set
            {
                dept = value;
            }
        }

        public string Date
        {
            get
            {
                return date;
            }

            set
            {
                date = value;
            }
        }

        public object BgColor
        {
            get
            {
                return bgColor;
            }

            set
            {
                bgColor = value;
            }
        }
        private int getWorkOrRest() {
            //16777215: 白色。
            //32768：绿色。
            ///为绿色时,返回休息标志.
            if (bgColor.ToString() == "32768") {
                return 0;
            }
            return 1;
        }
        #region 更新方法.
        public int updateWorkSchedule() {
            int affectedCount = 0;
            string sqlStr = String.Format(@"
                                            UPDATE Work_Schedule
                                            SET Work_OR_Rest ={2} ,
                                                RECORD_TIME = SYSDATE
                                            WHERE dept = '{0}'
                                            AND Work_And_Rest_Date = TO_DATE('{1}','YYYY-MM-DD') 
                                            ", this.dept,
                                               this.date,
                                               this.getWorkOrRest()
                                               );
            affectedCount = Tools.OracleDaoHelper.executeSQL(sqlStr);
            return affectedCount;
        }
        #endregion
    }
}
