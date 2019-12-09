using Oracle.DataAccess.Client;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using Tools;
namespace AttendanceRecord.Entities
{
    /// <summary>
    /// 用来存储考勤记录
    /// </summary>
    public class AttendanceR
    {
       static string  _start_date; 
       static string  _end_date; 
       static string  _tabulation_time;
       static string _file_path;
       static  string _sheet_name;
       static char _prefix_Job_Number;
      
       string _fingerprint_Date;
       string _job_number; 
       string _name; 
       string _dept;
       static string _random_Str;
        //第一次按指纹时间点
       string _FPT_Fisrt_Time;
        //最后一次按指纹时间点
       string _FPT_Last_Time;
        OracleHelper oH = null;
      /// <summary>
        ///
        /// </summary>
        /// <param name="_start_date">起始日期</param>
        /// <param name="_end_date">终止日期</param>
        /// <param name="_tabulation_time">制表日期</param>
        public AttendanceR()
        {
        }
        public AttendanceR(string name) {
            this._name = name;
        }
        /*
        public string[] getJNByName() {
            string[] jNArray = { };
            string sqlStr = string.Format(@"SELECT DISTINCT JOB_NUMBER 
                                              FROM Attendance_Record AR 
                                              WHERE Name = '{0}'
                                              AND trunc(FingerPrint_Date,'MM') >= TRUNC(ADD_MONTHS(SYSDATE,-2),'MM')", this._name);
            DataTable dt = OracleDaoHelper.getDTBySql(sqlStr);
            if (dt.Rows.Count==0) {
                return jNArray;
            }
            jNArray = new String[dt.Rows.Count];
            for (int i=0;i<=dt.Rows.Count-1;i++) {
                jNArray[i] = dt.Rows[i]["JOB_NUMBER"].ToString();
            }
            return jNArray;
        }
        */
        public string Fingerprint_Date
        {
            get
            {
                return _fingerprint_Date;
            }

            set
            {
                _fingerprint_Date = value;
            }
        }
        public string Job_number
        {
            get
            {
                return _job_number;
            }

            set
            {
                _job_number = value;
            }
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
        public string Dept
        {
            get
            {
                return _dept;
            }

            set
            {
                _dept = value;
            }
        }
       
        public static string Start_date
        {
            get
            {
                return _start_date;
            }

            set
            {
                _start_date = value;
            }
        }

        public static string End_date
        {
            get
            {
                return _end_date;
            }

            set
            {
                _end_date = value;
            }
        }

        public static string Tabulation_time
        {
            get
            {
                return _tabulation_time;
            }

            set
            {
                _tabulation_time = value;
            }
        }

        public string FPT_Fisrt_Time
        {
            get
            {
                return _FPT_Fisrt_Time;
            }

            set
            {
                _FPT_Fisrt_Time = value;
            }
        }

        public string FPT_Last_Time
        {
            get
            {
                return _FPT_Last_Time;
            }

            set
            {
                _FPT_Last_Time = value;
            }
        }

        public static string Sheet_name
        {
            get
            {
                return _sheet_name;
            }

            set
            {
                _sheet_name = value;
            }
        }

        public static string Random_Str
        {
            get
            {
                return _random_Str;
            }

            set
            {
                _random_Str = value;
            }
        }

        public static char Prefix_Job_Number
        {
            get
            {
                return _prefix_Job_Number;
            }

            set
            {
                _prefix_Job_Number = value;
            }
        }

        public static string File_path
        {
            get
            {
                return _file_path;
            }

            set
            {
                _file_path = value;
            }
        }
        /*
        /// <summary>
        ///用于事务提交.
        /// </summary>
        /// <param name="conn"></param>
        /// <returns></returns>
        public int saveDataToDB(Oracle.DataAccess.Client.OracleConnection conn) {
            int affectedCount = 0;
            string sqlStr;
            if (!ifExistsFPRecord()) {
                sqlStr = String.Format(@"INSERT INTO Attendance_Record(
                                                 SEQ,
                                                 Start_Date,
                                                 End_Date,
                                                 Tabulation_Time,
                                                 FingerPrint_Date,
                                                 Job_Number,
                                                 Name,
                                                 Dept,
                                                 Sheet_Name,
                                                 FPT_First_Time,
                                                 FPT_Last_Time,
                                                Random_Str
                                          )
                                          VALUES(
                                                 s_attendancerecord.Nextval,
                                                 TO_DATE('{0}','yyyy-MM-dd'),
                                                TO_DATE('{1}','yyyy-MM-dd'),
                                                TO_DATE('{2}','yyyy-MM-dd'),
                                                 TO_DATE('{3}','yyyy-MM-dd'),
                                                 '{4}',
                                                 '{5}',
                                                 '{6}',
                                                 '{7}',
                                                  TO_DATE('{8}','yyyy-MM-dd HH24:MI'),
                                                 TO_DATE('{9}','yyyy-MM-dd HH24:MI'),
                                                  '{10}'
                                          )",
                                         AttendanceR.Start_date,
                                         AttendanceR.End_date,
                                         AttendanceR.Tabulation_time,
                                         this.Fingerprint_Date,
                                         this.Job_number,
                                         this.Name,
                                         this.Dept,
                                         AttendanceR.Sheet_name,
                                         this.FPT_Fisrt_Time,
                                         this.FPT_Last_Time,
                                         AttendanceR.Random_Str
                                         );
                affectedCount = Tools.OracleDaoHelper.executeSQLBySpecificConn(sqlStr, conn);
                return affectedCount;
            }
            sqlStr = String.Format(@"UPDATE Attendance_Record
                                          SET Start_Date =TO_DATE('{3}','yyyy-MM-dd'),
                                              END_DATE = TO_DATE('{4}','yyyy-MM-dd'),
                                              TABULATION_TIME =  TO_DATE('{5}','yyyy-MM-dd'),
                                              DEPT = '{6}',
                                              FPT_First_TIME = TO_DATE('{7}','yyyy-MM-dd HH24:MI'),
                                              FPT_LAST_TIME = TO_DATE('{8}','yyyy-MM-dd HH24:MI'),
                                              COME_LATE_NUM = 0,
                                              LEAVE_EARLY_NUM = 0,
                                               Sheet_Name = '{9}',
                                                Random_Str = '{10}'
                                            WHERE JOB_NUMBER = '{0}'
                                            AND NAME = '{1}'
                                            AND FingerPrint_Date = TO_DATE('{2}','yyyy-MM-dd')",
                                          this.Job_number,
                                          this.Name,
                                          this.Fingerprint_Date,
                                          AttendanceR.Start_date,
                                          AttendanceR.End_date,
                                          AttendanceR._tabulation_time,
                                          this.Dept,
                                          this.FPT_Fisrt_Time,
                                          this.FPT_Last_Time,
                                          AttendanceR.Sheet_name,
                                          AttendanceR.Random_Str
                                          );
            return Tools.OracleDaoHelper.executeSQLBySpecificConn(sqlStr, conn) ;
        }
        */
        /// <summary>
        ///用于事务提交.
        /// </summary>
        /// <param name="conn"></param>
        /// <returns></returns>
        public int saveDataToDB()
        {
            string procedureName = "save_ar_func";
            OracleParameter param_jN = new OracleParameter("v_job_number", OracleDbType.Varchar2, ParameterDirection.Input);
            OracleParameter param_Name = new OracleParameter("v_name", OracleDbType.NVarchar2, ParameterDirection.Input);
            OracleParameter param_Dept = new OracleParameter("v_dept", OracleDbType.Varchar2, ParameterDirection.Input);
            OracleParameter param_Start_Date = new OracleParameter("v_start_date", OracleDbType.Varchar2, ParameterDirection.Input);
            OracleParameter param_End_Date = new OracleParameter("v_end_date", OracleDbType.Varchar2, ParameterDirection.Input);
            OracleParameter param_Tabulation_Time = new OracleParameter("v_tabulation_time", OracleDbType.Varchar2, ParameterDirection.Input);
            OracleParameter param_FingerPrint_Date = new OracleParameter("v_fingerPrint_date", OracleDbType.Varchar2, ParameterDirection.Input);
            OracleParameter param_fpt_first_Time = new OracleParameter("v_fpt_first_time", OracleDbType.Varchar2, ParameterDirection.Input);
            OracleParameter param_fpt_End_Time = new OracleParameter("v_fpt_end_time", OracleDbType.Varchar2, ParameterDirection.Input);
            OracleParameter param_file_path = new OracleParameter("v_file_path", OracleDbType.NVarchar2, ParameterDirection.Input);
            OracleParameter param_fpt_Sheet_Name = new OracleParameter("v_sheet_name", OracleDbType.NVarchar2, ParameterDirection.Input);
            OracleParameter param_random_Str = new OracleParameter("v_random_str", OracleDbType.Varchar2, ParameterDirection.Input);
            OracleParameter param_affected_count = new OracleParameter("v_affected_count", OracleDbType.Int32, ParameterDirection.ReturnValue);
            param_jN.Value = this._job_number;
            param_Name.Value = this._name;
            param_Dept.Value = this._dept;
            param_Start_Date.Value = AttendanceR._start_date;
            param_End_Date.Value = AttendanceR._end_date;
            param_Tabulation_Time.Value = AttendanceR._tabulation_time;
            param_FingerPrint_Date.Value = this._fingerprint_Date;
            param_fpt_first_Time.Value = this._FPT_Fisrt_Time;
            param_fpt_End_Time.Value = this._FPT_Last_Time;
            param_file_path.Value = AttendanceR.File_path;
            param_fpt_Sheet_Name.Value = AttendanceR._sheet_name;
            param_random_Str.Value = AttendanceR._random_Str;
            OracleParameter[] parameters = new OracleParameter[13] {
                                                                    param_affected_count,
                                                                    param_jN,
                                                                    param_Name,
                                                                    param_Dept,
                                                                    param_Start_Date,
                                                                    param_End_Date,
                                                                    param_Tabulation_Time,
                                                                    param_FingerPrint_Date,
                                                                    param_fpt_first_Time,
                                                                    param_fpt_End_Time,
                                                                    param_file_path,
                                                                    param_fpt_Sheet_Name,
                                                                    param_random_Str
                                                                        };
            if (oH==null) {
                oH = OracleHelper.getBaseDao();
            }
            oH.ExecuteNonQuery(procedureName, parameters);
            int j=Int32.Parse(parameters[0].Value.ToString());
            return j;
        }
        #region 判断是否存在 某个员工的刷卡记录.
        public bool ifExistsFPRecord() {
            string sqlStr = String.Format(@"SELECT 1
                                              FROM Attendance_Record
                                              WHERE  JOB_NUMBER = '{0}'
                                              AND FingerPrint_Date = TO_DATE('{1}','yyyy-MM-dd')",
                                              Job_number,
                                              Fingerprint_Date);
            return Tools.OracleDaoHelper.getDTBySql(sqlStr).Rows.Count>0? true :false;
        }
        #endregion
        public System.Data.DataTable getAR(string random_Str)
        {
            string procedureName = "pkg_ar.get_ARDetail_By_Random_Str";
            OracleParameter param_random_Str = new OracleParameter("v_random_str", OracleDbType.Varchar2, ParameterDirection.Input);
            OracleParameter param_Cur_AR = new OracleParameter("v_cur_ar", OracleDbType.RefCursor, ParameterDirection.ReturnValue);
            param_random_Str.Value = random_Str;
            OracleParameter[] parameters = new OracleParameter[2] { param_Cur_AR,
                                                                        param_random_Str
                                                                   };
            if (oH == null)
            {
                oH = OracleHelper.getBaseDao();
            }
            DataTable dt = oH.getDT(procedureName, parameters);
            return dt;
        }
        /// <summary>
        /// 按日期查询。
        /// </summary>
        /// <param name="year_And_Month_Str"></param>
        /// <returns></returns>
        public System.Data.DataTable getARByYearAndMonth(string year_And_Month_Str)
        {
            String sqlStr = String.Format(@"SELECT 
                                              start_date AS ""起始日期"", 
                                              end_date AS ""终止日期"", 
                                              tabulation_time AS ""制表日期"", 
                                              finger_print_date AS ""指纹日期"", 
                                              dept AS ""部门"",
                                              JOB_NUMBER AS ""工号"", 
                                              name AS ""姓名"", 
                                              SUBSTR(TO_CHAR(fpt_first_time,'yyyy-MM-dd HH24:MI:SS'),12,5) AS ""上班时间点"", 
                                              SUBSTR(TO_CHAR(fpt_last_time,'yyyy-MM-dd HH24:MI:SS'),12,5) AS ""下班时间点"", 
                                              sheet_name As ""表格名称"", 
                                              reord_time AS ""记录时间"", 
                                              come_num AS ""出勤天数"",
                                              not_finger_print AS ""未打卡"", 
                                              delay_time AS ""延点"", 
                                              come_late_num AS ""迟到次数"",
                                              leave_early_num AS ""早退次数"", 
                                              dinner_subsidy AS ""餐补"", 
                                              random_str  AS ""随机字符串""
                              FROM Attendance_Record_Briefly
                              WHERE TRUNC(finger_print_date,'MM')= TO_DATE('{0}','yyyy-MM')
                              OrDER bY 
                                        JOB_NUMBER ASC,
                                        finger_print_date ASC
                                         ", year_And_Month_Str);
            System.Data.DataTable dt = OracleDaoHelper.getDTBySql(sqlStr);
            return dt;
        }
        /// <summary>
        /// 按日期范围查询考勤记录。
        /// </summary>
        /// <param name="startDateStr"></param>
        /// <param name="endDateStr"></param>
        /// <returns></returns>
        public System.Data.DataTable getARByRange(string startDateStr,string endDateStr)
        {
            String sqlStr = String.Format(@"SELECT 
                                              start_date AS ""起始日期"", 
                                              end_date AS ""终止日期"", 
                                              tabulation_time AS ""制表日期"", 
                                              finger_print_date AS ""指纹日期"", 
                                              dept AS ""部门"",
                                              JOB_NUMBER AS ""工号"", 
                                              name AS ""姓名"", 
                                              SUBSTR(TO_CHAR(fpt_first_time,'yyyy-MM-dd HH24:MI:SS'),12,5) AS ""上班时间点"", 
                                              SUBSTR(TO_CHAR(fpt_last_time,'yyyy-MM-dd HH24:MI:SS'),12,5) AS ""下班时间点"", 
                                              sheet_name As ""表格名称"", 
                                              reord_time AS ""记录时间"", 
                                              come_num AS ""出勤天数"",
                                              not_finger_print AS ""未打卡"", 
                                              delay_time AS ""延点"", 
                                              come_late_num AS ""迟到次数"",
                                              leave_early_num AS ""早退次数"", 
                                              dinner_subsidy AS ""餐补"", 
                                              random_str  AS ""随机字符串""
                              FROM Attendance_Record_Briefly
                              WHERE finger_print_date between  TO_DATE('{0}','yyyy-MM-dd') and TO_DATE('{1}','yyyy-MM-dd')
                              OrDER bY 
                                        JOB_NUMBER ASC,
                                        finger_print_date ASC
                                         ", startDateStr,endDateStr);
            System.Data.DataTable dt = OracleDaoHelper.getDTBySql(sqlStr);
            return dt;
        }
        /*
        #region 获取该日期范围内有多少日的考勤记录。
        /// <summary>
        /// 获取该日期范围内有多少日的考勤记录
        /// </summary>
        /// <param name="Year_And_Month_Str"></param>
        /// <returns></returns>
        public static int get_AR_Days_Num(string Year_And_Month_Str)
        {
            string sqlStr = string.Format(@"
                                            SELECT COUNT(1)
                                            FROM 
                                            (
                                               select AR.Fingerprint_Date
                                               from Attendance_Record AR
                                               where TRUNC(AR.Fingerprint_Date,'MM') = To_DATE('{0}','yyyy-MM')
                                               GROUP BY AR.Fingerprint_Date
                                            ) TEMP",
                                       Year_And_Month_Str);
            int result = 0;
            int.TryParse(OracleDaoHelper.getDTBySql(sqlStr).Rows[0][0].ToString(), out result);
            return result;
        }
        #endregion
        */
        #region 获取该日期范围内有多少日的考勤记录。
        /// <summary>
        /// 获取该日期范围内有多少日的考勤记录
        /// </summary>
        /// <param name="Year_And_Month_Str"></param>
        /// <returns></returns>
        public static int get_AR_Days_Num(string Year_And_Month_Str)
        {
            string sqlStr = string.Format(@"
                                            SELECT COUNT(1)
                                            FROM 
                                            (
                                               select ARBriefly.finger_print_date
                                               from Attendance_Record_Briefly ARBriefly
                                               where TRUNC(ARBriefly.finger_print_date,'MM') = To_DATE('{0}','yyyy-MM')
                                               GROUP BY ARBriefly.finger_print_date
                                            ) TEMP",
                                       Year_And_Month_Str);
            int result = 0;
            int.TryParse(OracleDaoHelper.getDTBySql(sqlStr).Rows[0][0].ToString(), out result);
            return result;
        }
        #endregion
        #region
        /// <summary>
        /// 获取该日期范围内有多少日的考勤记录
        /// </summary>
        /// <param name="Year_And_Month_Str"></param>
        /// <returns></returns>
        public static int get_A_R_Summary_Days_Num(string Year_And_Month_Str)
        {
            string sqlStr = string.Format(@"
                                            SELECT COUNT(1)
                                            FROM 
                                            (
                                               select A_R_Summary.Fingerprint_Date
                                               from Attendance_Record_Summary A_R_Summary
                                               where TRUNC(A_R_Summary.Fingerprint_Date,'MM') = To_DATE('{0}','yyyy-MM')
                                               GROUP BY A_R_Summary.Fingerprint_Date
                                            ) TEMP",
                                       Year_And_Month_Str);
            int result = 0;
            int.TryParse(OracleDaoHelper.getDTBySql(sqlStr).Rows[0][0].ToString(), out result);
            return result;
        }
        #endregion
        #region 获取该日期范围内有多少日的考勤记录。
        /// <summary>
        /// 
        /// </summary>
        /// <param name="Year_And_Month_Str"></param>
        /// <param name="prefix_Job_Number">工号前缀</param>
        /// <returns></returns>
        public static int get_AR_Days_Num(string Year_And_Month_Str,string prefix_Job_Number)
        {
            string sqlStr = string.Format(@"
                                            SELECT COUNT(1)
                                            FROM 
                                            (
                                               select ARBriefly.Finger_print_date
                                               from Attendance_Record_Briefly ARBriefly
                                               where TRUNC(ARBriefly.Finger_print_date,'MM') = To_DATE('{0}','yyyy-MM')
                                                and substr(ARBriefly.Job_Number,1,9) = '{1}'
                                               GROUP BY ARBriefly.Finger_print_date
                                            ) TEMP",
                                       Year_And_Month_Str, prefix_Job_Number);
            int result = 0;
            int.TryParse(OracleDaoHelper.getDTBySql(sqlStr).Rows[0][0].ToString(), out result);
            return result;
        }
        #endregion
        /// <summary>
        /// 统计有多少人。该月参加考勤的有多少人。
        /// </summary>
        /// <param name="Year_And_Month_Str"></param>
        /// <returns></returns>
        public static int get_Total_Num_Of_Staffs(string Year_And_Month_Str)
        {
            string sqlStr = string.Format(@"
                                            SELECT COUNT(1)
                                            FROM 
                                            (
                                                select AR.JOB_NUMBER
                                               from Attendance_Record AR
                                               where TRUNC(AR.Fingerprint_Date,'MM') = To_DATE('{0}','yyyy-MM')
                                               GROUP BY AR.JOB_NUMBER
                                            ) TEMP",
                                       Year_And_Month_Str);
            int result = 0;
            int.TryParse(OracleDaoHelper.getDTBySql(sqlStr).Rows[0][0].ToString(), out result);
            return result;
        }
        /// <summary>
        /// 获取该月该考勤机所统计的同名用户的个数。
        /// </summary>
        /// <param name="Year_And_Month_Str"></param>
        /// <param name="prefix_Job_Number"></param>
        /// <returns></returns>
        public static int get_Total_Num_Of_Staffs_By_YAndM_And_PreJN(string Year_And_Month_Str,string prefix_Job_Number)
        {
            string sqlStr = string.Format(@"
                                            SELECT 1
                                            FROM Attendance_Record_Briefly ARBriefly
                                            WHERE ARBriefly.NAME = ANY(
                                                    SELECT DISTINCT NAME 
                                                    FROM Attendance_Record 
                                                    WHERE TRUNC(Finger_print_Date,'MM') = To_DATE('{0}','yyyy-MM')
                                                    AND substr(Job_Number,1,9) = '{1}'
                                                )
                                            AND TRUNC(ARBriefly.Finger_print_Date,'MM') = To_DATE('{0}','yyyy-MM')
                                            GROUP BY AR.JOB_NUMBER
                                            ",
                                       Year_And_Month_Str, prefix_Job_Number);
            return OracleDaoHelper.getDTBySql(sqlStr).Rows.Count;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="Year_And_Month_Str"></param>
        /// <param name="prefix_Job_Number">工号前缀</param>
        /// <returns></returns>
        public static int get_Total_Num_Of_Staffs_By_YAndM(string Year_And_Month_Str)
        {
            string sqlStr = string.Format(@"
                                            SELECT COUNT(1)
                                            FROM 
                                            (
                                                select ARBriefly.JOB_NUMBER
                                               from Attendance_Record_Briefly ARBriefly
                                               where TRUNC(ARBriefly.Finger_print_Date,'MM') = To_DATE('{0}','yyyy-MM')
                                               GROUP BY ARBriefly.JOB_NUMBER
                                            ) TEMP",
                                       Year_And_Month_Str );
            int result = 0;
            int.TryParse(OracleDaoHelper.getDTBySql(sqlStr).Rows[0][0].ToString(), out result);
            return result;
        }
        /*
        /// <summary>
        /// 统计有多少人，此月此考勤机。
        /// </summary>
        /// <param name="Year_And_Month_Str"></param>
        /// <param name="prefix_Job_Number">工号前缀</param>
        /// <returns></returns>
        public static int get_Total_Num_Of_Staffs_By_YAndM_And_AMFlag(string  attendance_machine_flag, string Year_And_Month_Str )
        {
            string sqlStr = string.Format(@"
                                            SELECT COUNT(1)
                                            FROM 
                                            (
                                                select AR.JOB_NUMBER
                                               from Attendance_Record AR
                                               where substr(job_number,1,1) in  ({0})
                                               AND TRUNC(AR.Fingerprint_Date,'MM') = To_DATE('{1}','yyyy-MM')
                                               GROUP BY AR.JOB_NUMBER
                                            ) TEMP",
                                            attendance_machine_flag,
                                            Year_And_Month_Str
                                            );
            int result = 0; 
            int.TryParse(OracleDaoHelper.getDTBySql(sqlStr).Rows[0][0].ToString(), out result);
            return result;
        }
        */
        /// <summary>
        /// 统计有多少人，此月此考勤机。
        /// </summary>
        /// <param name="Year_And_Month_Str"></param>
        /// <param name="prefix_Job_Number">工号前缀</param>
        /// <returns></returns>
        public static int get_Total_Num_Of_Staffs_By_YAndM_And_AMFlag(string attendance_machine_flag, string Year_And_Month_Str)
        {
            string sqlStr = string.Format(@"
                                            SELECT COUNT(1)
                                            FROM 
                                            (
                                                select ARBriefly.JOB_NUMBER
                                               from Attendance_Record_Briefly ARBriefly
                                               where substr(job_number,1,1) in  ({0})
                                               AND TRUNC(ARBriefly.Finger_print_Date,'MM') = To_DATE('{1}','yyyy-MM')
                                               GROUP BY ARBriefly.JOB_NUMBER
                                            ) TEMP",
                                            attendance_machine_flag,
                                            Year_And_Month_Str
                                            );
            int result = 0;
            int.TryParse(OracleDaoHelper.getDTBySql(sqlStr).Rows[0][0].ToString(), out result);
            return result;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="Year_And_Month_Str"></param>
        /// <param name="prefix_Job_Number">工号前缀</param>
        /// <returns></returns>
        public static int get_Total_Num_Of_Staffs_Of_A_R_Summary_By_YAndM(string Year_And_Month_Str)
        {
            string sqlStr = string.Format(@"
                                            SELECT COUNT(1)
                                            FROM 
                                            (
                                                select A_R_Summary.JOB_NUMBER
                                               from Attendance_Record_Summary A_R_Summary
                                               where TRUNC(A_R_Summary.Fingerprint_Date,'MM') = To_DATE('{0}','yyyy-MM')
                                               GROUP BY A_R_Summary.JOB_NUMBER
                                            ) TEMP",
                                       Year_And_Month_Str);
            int result = 0;
            int.TryParse(OracleDaoHelper.getDTBySql(sqlStr).Rows[0][0].ToString(), out result);
            return result;
        }
        /// <summary>
        /// 
        /// </summary>
        public void combine_Job_Number() {
            this._job_number = this.Job_number.PadLeft(3, '0').PadLeft(12,AttendanceR.Prefix_Job_Number); 
        }

        public DataTable getARNameLastThreeMonth(string lastName)
        {
            string sqlstr = string.Format(@"select distinct ar.name
                                            from attendance_record ar
                                            where ar.name like '{0}%'
                                            and ar.fingerprint_date >= trunc(add_months(sysdate,-2),'MM')
                                            ORDER BY NLSSORT(name, 'NLS_SORT = SCHINESE_PINYIN_M') asc"
                                            ,lastName);
            return OracleDaoHelper.getDTBySql(sqlstr);
        }
        public DataTable getARNameLastThreeMonth()
        {
            string sqlstr = string.Format(@"select distinct ar.name
                                            from attendance_record ar
                                            where ar.name like '{0}%'
                                            and ar.fingerprint_date >= trunc(add_months(sysdate,-2),'MM')
                                            ORDER BY NLSSORT(name, 'NLS_SORT = SCHINESE_PINYIN_M') asc"
                                            , this._name);
            return OracleDaoHelper.getDTBySql(sqlstr);
        }
        /*
        string start_date_str,
                                            string end_date_str,
                                            string tabulation_time_str,
                                            string fingerprint_date_str,
                                            string job_number,
                                            string name,
                                            string dept,
                                            string fpt_first_time,
                                            string fpt_last_time,
                                            string filePath
            */
        /// <summary>
        /// 此方法用于姓名相同时存入该表格。
        /// </summary>
        /// <param name="start_date_str"></param>
        /// <param name="end_date_str"></param>
        /// <param name="tabulation_time_str"></param>
        /// <param name="fingerprint_date_str"></param>
        /// <param name="job_number"></param>
        /// <param name="name"></param>
        /// <param name="dept"></param>
        /// <param name="fpt_first_time"></param>
        /// <param name="fpt_last_time"></param>
        ///
        public int importAR(OracleConnection conn)
        {
            string procName = "PKG_Import_AR.import_AR";
            OracleParameter param_affected_count = new OracleParameter("v_affected_count", OracleDbType.Int32, ParameterDirection.ReturnValue);
            OracleParameter param_start_date = new OracleParameter("v_start_date_str", OracleDbType.Varchar2, System.Data.ParameterDirection.Input);
            OracleParameter param_end_date = new OracleParameter("v_end_date_str", OracleDbType.Varchar2, ParameterDirection.Input);
            OracleParameter param_tabulation_time = new OracleParameter("v_tabulation_time_str", OracleDbType.Varchar2, ParameterDirection.Input);
            OracleParameter param_fingerprint_date = new OracleParameter("v_fingerprint_date_str", OracleDbType.Varchar2, ParameterDirection.Input);
            OracleParameter param_jN = new OracleParameter("v_job_number", OracleDbType.Varchar2, ParameterDirection.Input);
            OracleParameter param_name = new OracleParameter("v_name", OracleDbType.NVarchar2, ParameterDirection.Input);
            OracleParameter param_dept = new OracleParameter("v_dept", OracleDbType.NVarchar2, ParameterDirection.Input);
            OracleParameter param_fpt_first_time = new OracleParameter("v_fpt_first_time", OracleDbType.Varchar2, ParameterDirection.Input);
            OracleParameter param_fpt_last_time = new OracleParameter("v_fpt_last_time", OracleDbType.Varchar2, ParameterDirection.Input);
            OracleParameter param_file_path = new OracleParameter("v_file_path", OracleDbType.NVarchar2, ParameterDirection.Input);
            OracleParameter param_random_str = new OracleParameter("v_random_str", OracleDbType.Varchar2, ParameterDirection.Input);

            param_start_date.Value = AttendanceR._start_date;
            param_end_date.Value = AttendanceR._end_date;
            param_tabulation_time.Value = AttendanceR._tabulation_time;
            param_fingerprint_date.Value = this._fingerprint_Date;
            param_jN.Value = this._job_number;
            param_name.Value = this._name;
            param_dept.Value = this._dept;
            param_fpt_first_time.Value = this._FPT_Fisrt_Time;
            param_fpt_last_time.Value = this._FPT_Last_Time;
            param_file_path.Value = AttendanceR._file_path;
            param_random_str.Value = AttendanceR._random_Str;
            OracleParameter[] paramters = new OracleParameter[12] {
                                     param_affected_count,
                                     param_start_date,
                                     param_end_date,
                                     param_tabulation_time,
                                     param_fingerprint_date,
                                     param_jN,
                                     param_name,
                                     param_dept,
                                     param_fpt_first_time,
                                     param_fpt_last_time,
                                     param_file_path,
                                     param_random_str
            };
            OracleHelper oH = OracleHelper.getBaseDao();
            oH.ExecuteSPWithSpecificConn(procName, paramters,conn);
            return int.Parse(paramters[0].Value.ToString());
        }
        /// <summary>
        /// 此方法用于姓名相同时存入该表格。
        /// </summary>
        /// <param name="start_date_str"></param>
        /// <param name="end_date_str"></param>
        /// <param name="tabulation_time_str"></param>
        /// <param name="fingerprint_date_str"></param>
        /// <param name="job_number"></param>
        /// <param name="name"></param>
        /// <param name="dept"></param>
        /// <param name="fpt_first_time"></param>
        /// <param name="fpt_last_time"></param>
        ///
        public int import_AR_To_Preparative_Table()
        {
            string procName = "PKG_Import_AR_To_P_Table.import_AR_To_Preparative_Table";

            OracleParameter param_affected_count = new OracleParameter("v_affected_count", OracleDbType.Int32, ParameterDirection.ReturnValue);
            OracleParameter param_start_date = new OracleParameter("v_start_date_str", OracleDbType.Varchar2, System.Data.ParameterDirection.Input);
            OracleParameter param_end_date = new OracleParameter("v_end_date_str", OracleDbType.Varchar2, ParameterDirection.Input);
            OracleParameter param_tabulation_time = new OracleParameter("v_tabulation_time_str", OracleDbType.Varchar2, ParameterDirection.Input);
            OracleParameter param_fingerprint_date = new OracleParameter("v_fingerprint_date_str", OracleDbType.Varchar2, ParameterDirection.Input);
            OracleParameter param_jN = new OracleParameter("v_job_number", OracleDbType.Varchar2, ParameterDirection.Input);
            OracleParameter param_name = new OracleParameter("v_name", OracleDbType.NVarchar2, ParameterDirection.Input);
            OracleParameter param_dept = new OracleParameter("v_dept", OracleDbType.NVarchar2, ParameterDirection.Input);
            OracleParameter param_fpt_first_time = new OracleParameter("v_fpt_first_time", OracleDbType.Varchar2, ParameterDirection.Input);
            OracleParameter param_fpt_last_time = new OracleParameter("v_fpt_last_time", OracleDbType.Varchar2, ParameterDirection.Input);
            OracleParameter param_file_path = new OracleParameter("v_file_path", OracleDbType.NVarchar2, ParameterDirection.Input);
            OracleParameter param_random_str = new OracleParameter("v_random_str", OracleDbType.Varchar2, ParameterDirection.Input);

            param_start_date.Value = AttendanceR._start_date;
            param_end_date.Value = AttendanceR._end_date;
            param_tabulation_time.Value = AttendanceR._tabulation_time;
            param_fingerprint_date.Value = this._fingerprint_Date;
            param_jN.Value = this._job_number;
            param_name.Value = this._name;
            param_dept.Value = this._dept;
            param_fpt_first_time.Value = this._FPT_Fisrt_Time;
            param_fpt_last_time.Value = this._FPT_Last_Time;
            param_file_path.Value = AttendanceR._file_path;
            param_random_str.Value = AttendanceR._random_Str;
            OracleParameter[] paramters = new OracleParameter[12] {
                                     param_affected_count,
                                     param_start_date,
                                     param_end_date,
                                     param_tabulation_time,
                                     param_fingerprint_date,
                                     param_jN,
                                     param_name,
                                     param_dept,
                                     param_fpt_first_time,
                                     param_fpt_last_time,
                                     param_file_path,
                                     param_random_str
            };
            OracleHelper oH = OracleHelper.getBaseDao();
            oH.ExecuteNonQuery(procName, paramters);
            return int.Parse(paramters[0].Value.ToString());
        }
    }
}
