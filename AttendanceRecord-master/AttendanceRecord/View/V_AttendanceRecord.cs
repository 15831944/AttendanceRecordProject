using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using Tools;
namespace AttendanceRecord.View
{
    public class V_AttendanceRecord
    {
        public class AR_Properties {
            private string _start_Date;
            private string _end_Date;
            private string _tabulation_date;

            public string Start_Date
            {
                get
                {
                    return _start_Date;
                }

                set
                {
                    _start_Date = value;
                }
            }

            public string End_Date
            {
                get
                {
                    return _end_Date;
                }

                set
                {
                    _end_Date = value;
                }
            }

            public string Tabulation_date
            {
                get
                {
                    return _tabulation_date;
                }

                set
                {
                    _tabulation_date = value;
                }
            }

            public AR_Properties(string _start_Date, string _end_Date, string _tabulation_date)
            {
                this.Start_Date = _start_Date;
                this.End_Date = _end_Date;
                this.Tabulation_date = _tabulation_date;
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="Year_And_Month_Str"></param>
        /// <param name="prefix_Job_Number"></param>
        /// <returns></returns>
        #region 获取该月该考勤机的出勤人数。
        public static int getARNum(string Year_And_Month_Str, string prefix_Job_Number) {
            string sqlStr = string.Format(@"SELECT 1
                                              FROM Attendance_Record AR
                                              WHERE TRUNC(AR.Fingerprint_Date,'MM') = TO_DATE('{1}','YYYY-MM')
                                                AND Name = ANY(
                                                        SELECT DISTINCT NAME 
                                                        FROM Attendance_Record 
                                                        WHERE SUBSTR(Job_Number,1,9) = '{0}'
                                                        AND TRUNC(Fingerprint_Date,'MM') = TO_DATE('{1}','YYYY-MM')
                                                    )    
                                                GROUP BY AR.JOB_NUMBER",
                                                   prefix_Job_Number,
                                                   Year_And_Month_Str);
            return OracleDaoHelper.getDTBySql(sqlStr).Rows.Count;
        }
        #endregion
        #region 获取该月该考勤机的出勤人数。
        public static int getARNumByYearAndMonth(string Year_And_Month_Str)
        {
            string sqlStr = string.Format(@"SELECT 1
                                              FROM Attendance_Record AR
                                              WHERE  TRUNC(AR.Fingerprint_Date,'MM') = TO_DATE('{0}','YYYY-MM')
                                                    GROUP BY AR.JOB_NUMBER",
                                                   Year_And_Month_Str);
            return OracleDaoHelper.getDTBySql(sqlStr).Rows.Count;
        }
        #endregion
        /*
        #region 获取该月该考勤机的出勤人数。
        public static int getARNumByAMFlag_YearAndMonth(string Attendance_Machine_Flag,string Year_And_Month_Str)
        {
            string sqlStr = string.Format(@"SELECT 1
                                              FROM Attendance_Record AR
                                              WHERE SubStr(job_number,1,1) in ({0})
                                                AND TRUNC(AR.Fingerprint_Date,'MM') = TO_DATE('{1}','YYYY-MM')
                                                    GROUP BY AR.JOB_NUMBER",
                                                   Attendance_Machine_Flag,
                                                   Year_And_Month_Str);
            return OracleDaoHelper.getDTBySql(sqlStr).Rows.Count;
        }
        #endregion
        */
        #region 获取该月该考勤机的出勤人数。
        public static int getARNumByAMFlag_YearAndMonth(string Attendance_Machine_Flag, string Year_And_Month_Str)
        {
            string sqlStr = string.Format(@"SELECT 1
                                              FROM Attendance_Record_Briefly ARBriefly
                                              WHERE SubStr(job_number,1,1) in ({0})
                                                AND TRUNC(ARBriefly.Finger_print_Date,'MM') = TO_DATE('{1}','YYYY-MM')
                                                    GROUP BY ARBriefly.JOB_NUMBER",
                                                   Attendance_Machine_Flag,
                                                   Year_And_Month_Str);
            return OracleDaoHelper.getDTBySql(sqlStr).Rows.Count;
        }
        #endregion
        #region 获取该月该考勤机的出勤人数。  从Attendance_Record_Summary中获取
        public static int getARSummaryNumByYearAndMonth(string Year_And_Month_Str)
        {
            string sqlStr = string.Format(@"SELECT 1
                                              FROM Attendance_Record_Summary ARSummary
                                              WHERE  TRUNC(ARSummary.Fingerprint_Date,'MM') = TO_DATE('{0}','YYYY-MM')
                                                    GROUP BY ARSummary.JOB_NUMBER",
                                                   Year_And_Month_Str);
            return OracleDaoHelper.getDTBySql(sqlStr).Rows.Count;
        }
        #endregion
        #region 获取该月某考勤记录的属性。
        public static AR_Properties getARProperties(string Year_And_Month_Str, string prefix_Job_Number) {
            string sqlStr = string.Format(@"SELECT Temp.Start_Date,
                                                    Temp.End_Date,
                                                    Temp.Tabulation_Date
                                            FROM 
                                            (SELECT TO_CHAR(AR.start_Date,'YYYY-MM-DD') Start_Date,
                                                    TO_CHAR(AR.end_Date,'YYYY-MM-DD') End_Date,
                                                    TO_Char(AR.tabulation_Time,'YYYY-MM-DD') Tabulation_Date,
                                                    rowNum row_num
                                              FROM Attendance_Record AR
                                              WHERE SUBSTR(AR.Job_Number,1,9) = '{0}'
                                                   AND TRUNC(AR.Fingerprint_Date,'MM') = TO_DATE('{1}','YYYY-MM')
                                                ) TEMP
                                                WHERE Temp.row_num = 1
                                                ",
                                                prefix_Job_Number,
                                                Year_And_Month_Str);
            DataTable dt = OracleDaoHelper.getDTBySql(sqlStr);
            return new AR_Properties(dt.Rows[0]["Start_Date"].ToString(), dt.Rows[0]["End_Date"].ToString(), dt.Rows[0]["Tabulation_Date"].ToString());
        }
        #endregion
        /*
        #region 获取该月某考勤记录的属性。
        public static AR_Properties getARProperties(string Year_And_Month_Str)
        {
            string sqlStr = string.Format(@"SELECT Temp.Start_Date,
                                                    Temp.End_Date,
                                                    Temp.Tabulation_Date
                                            FROM 
                                            (SELECT TO_CHAR(AR.start_Date,'YYYY-MM-DD') Start_Date,
                                                    TO_CHAR(AR.end_Date,'YYYY-MM-DD') End_Date,
                                                    TO_Char(AR.tabulation_Time,'YYYY-MM-DD') Tabulation_Date,
                                                    rowNum row_num
                                              FROM Attendance_Record AR
                                              WHERE TRUNC(AR.Fingerprint_Date,'MM') = TO_DATE('{0}','YYYY-MM')
                                                ) TEMP
                                                WHERE Temp.row_num = 1",
                                                Year_And_Month_Str);
            DataTable dt = OracleDaoHelper.getDTBySql(sqlStr);
            return new AR_Properties(dt.Rows[0]["Start_Date"].ToString(), dt.Rows[0]["End_Date"].ToString(), dt.Rows[0]["Tabulation_Date"].ToString());
        }
        #endregion
        */
        #region 获取该月某考勤记录的属性。
        public static AR_Properties getARProperties(string Year_And_Month_Str)
        {
            string sqlStr = string.Format(@"SELECT Temp.Start_Date,
                                                    Temp.End_Date,
                                                    Temp.Tabulation_Date
                                            FROM 
                                            (SELECT TO_CHAR(ARBriefly.start_Date,'YYYY-MM-DD') Start_Date,
                                                    TO_CHAR(ARBriefly.end_Date,'YYYY-MM-DD') End_Date,
                                                    TO_Char(ARBriefly.tabulation_Time,'YYYY-MM-DD') Tabulation_Date,
                                                    rowNum row_num
                                              FROM Attendance_Record_Briefly ARBriefly
                                              WHERE TRUNC(ARBriefly.Finger_print_Date,'MM') = TO_DATE('{0}','YYYY-MM')
                                                ) TEMP
                                                WHERE Temp.row_num = 1",
                                                Year_And_Month_Str);
            DataTable dt = OracleDaoHelper.getDTBySql(sqlStr);
            return new AR_Properties(dt.Rows[0]["Start_Date"].ToString(), dt.Rows[0]["End_Date"].ToString(), dt.Rows[0]["Tabulation_Date"].ToString());
        }
        #endregion
        #region 获取该月某考勤记录的属性。
        public static AR_Properties getPropertiesFromAttendanceRecordSummary(string Year_And_Month_Str)
        {
            string sqlStr = string.Format(@"SELECT Temp.Start_Date,
                                                    Temp.End_Date,
                                                    Temp.Tabulation_Date
                                            FROM 
                                            (SELECT TO_CHAR(ARS.start_Date,'YYYY-MM-DD') Start_Date,
                                                    TO_CHAR(ARS.end_Date,'YYYY-MM-DD') End_Date,
                                                    TO_Char(ARS.tabulation_Time,'YYYY-MM-DD') Tabulation_Date,
                                                    rowNum row_num
                                              FROM Attendance_Record_Summary ARS
                                              WHERE TRUNC(ARS.Fingerprint_Date,'MM') = TO_DATE('{0}','YYYY-MM')
                                                ) TEMP
                                                WHERE Temp.row_num = 1",
                                                Year_And_Month_Str);
            DataTable dt = OracleDaoHelper.getDTBySql(sqlStr);
            return new AR_Properties(dt.Rows[0]["Start_Date"].ToString(), dt.Rows[0]["End_Date"].ToString(), dt.Rows[0]["Tabulation_Date"].ToString());
        }
        #endregion

    }
}
