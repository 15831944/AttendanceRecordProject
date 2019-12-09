using Oracle.DataAccess.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Tools;
namespace AttendanceRecord.Entities
{
    public class AttendanceRecordDetail
    {
        public static  string _start_date;
        public static  string _end_date;
        public static  string _tabulation_time;
        public static char _prefix_Job_Number;
        private  string _fingerprint_date;
        private string _job_number;
        private string _name;
        private string _dept;
        public static string _sheet_name;
        private string _finger_print_time;

        private string _record_time;
        public static  string _random_str;
        public static  string _file_path;
        /// <summary>
        /// 用于填充考勤记录的标识。
        /// </summary>
        public string Fingerprint_date { get => _fingerprint_date; set => _fingerprint_date = value; }
        public string Job_number { get => _job_number; set => _job_number = value; }
        public string Name { get => _name; set => _name = value; }
        public string Dept { get => _dept; set => _dept = value; }
        public string Sheet_name { get => _sheet_name; set => _sheet_name = value; }
        public string Finger_print_time { get => _finger_print_time; set => _finger_print_time = value; }
        public void combine_Job_Number()
        {
            this._job_number = this.Job_number.PadLeft(3, '0').PadLeft(12, _prefix_Job_Number);
        }
        public bool ifExistsNullRecordOffingerPrintTime(string name,string year_and_month_str) {
            string sqlStr = string.Format(@"SELECT 1 FROM dual");
            return false;
        
        
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="conn"></param>
        public int saveBySpecificConn(OracleConnection conn) {
            string sqlStr = string.Format(@"  insert into attendance_record_detail(
                                                          start_date, 
                                                          end_date, 
                                                          tabulation_time, 
                                                          finger_print_date, 
                                                          job_number, 
                                                          name, 
                                                          dept, 
                                                          sheet_name, 
                                                          finger_print_time, 
                                                          random_str, 
                                                          file_path
                                                     )
                                               values(
                                                     to_date('{0}','yyyy-MM-dd'),               --<start_date>
                                                     to_date('{1}','yyyy-MM-dd'),           --<end_date>
                                                     to_Date('{2}','yyyy-MM-dd'),           --<tabulation_time>
                                                     to_date('{3}','yyyy-MM-dd'),           --<finger_print_date>
                                                     '{4}',                                 --<job_number>
                                                     '{5}',                                 --<name>
                                                     '{6}',                                 --<dept>
                                                     '{7}',                                 --<sheet_name>
                                                     to_date('{8}','yyyy-MM-dd HH24:MI'),        --<finger_print_time>
                                                     '{9}',                                    --<random_str>
                                                     '{10}'                                     --<file_path>
                                               )", 
                                                    _start_date,
                                                    _end_date,
                                                    _tabulation_time,
                                               this._fingerprint_date,
                                               this._job_number,
                                               this._name,
                                               this._dept,
                                                    _sheet_name,
                                               this._finger_print_time,
                                               _random_str,
                                               _file_path);
            return OracleDaoHelper.executeSQLBySpecificConn(sqlStr, conn);
        }
    }
}
