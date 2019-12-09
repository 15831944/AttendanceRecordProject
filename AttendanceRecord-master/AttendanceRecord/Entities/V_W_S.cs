using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Tools;
namespace AttendanceRecord.Entities
{
    public class V_W_S
    {
        private string dept; 
        private string work_and_rest_date;
        private string work_rate;
        private string work_or_rest;
        private string day_of_week;


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

        public string Work_and_rest_date
        {
            get
            {
                return work_and_rest_date;
            }

            set
            {
                work_and_rest_date = value;
            }
        }

        public string Work_rate
        {
            get
            {
                return work_rate;
            }

            set
            {
                work_rate = value;
            }
        }

        public string Work_or_rest
        {
            get
            {
                return work_or_rest;
            }

            set
            {
                work_or_rest = value;
            }
        }

        public string Day_of_week
        {
            get
            {
                return day_of_week;
            }

            set
            {
                day_of_week = value;
            }
        }
        #region 依据部门,日期获取该对象.
        public V_W_S get_V_W_S_By_Date_And_Dept() {
            V_W_S v_W_S = null;
            string sqlStr = String.Format(@"SELECT DEPT,
                                                    TO_CHAR(Work_And_Rest_Date,'YYYY-MM-DD') AS Work_And_Rest_Date,
                                                    CAST(Work_Rate AS VARCHAR2(10)) AS Work_Rate,
                                                    Work_OR_REST,
                                                    Day_Of_Week
                                            FROM V_W_S 
                                            WHERE Dept= '{0}'
                                                AND Work_And_Rest_Date = TO_DATE('{1}','YYYY-MM-DD')", this.dept,this.work_and_rest_date);
            v_W_S =ConvertHelper<V_W_S>.ConvertToList(OracleDaoHelper.getDTBySql(sqlStr))[0];
            return v_W_S;
        }
        #endregion
    }
}
