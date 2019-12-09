using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using Oracle.DataAccess.Client;
using Tools;
namespace AttendanceRecord.GetARSummary
{
    public class GetARSummary
    {
        /// <summary>
        /// 获取某个月，某考勤机的  考勤日期天数。
        /// </summary>
        /// <param name="v_attendance_machine_NO"></param>
        /// <param name="v_year_month_str"></param>
        /// <returns></returns>
        public static int get_nums_of_staffs(int v_attendance_machine_NO, string v_year_month_str)
        {
            OracleHelper oH = OracleHelper.getBaseDao();
            string proc_Name = "get_AR_Summary.get_nums_of_staffs";
            OracleParameter v_num = new OracleParameter("v_num", OracleDbType.Int32, ParameterDirection.ReturnValue);
            OracleParameter param_attendance_no = new OracleParameter("v_attendance_machine_no", OracleDbType.Varchar2, ParameterDirection.Input);
            OracleParameter param_year_month_str = new OracleParameter("v_year_month_str", OracleDbType.Varchar2, ParameterDirection.Input);
            param_year_month_str.Value = v_year_month_str;
            param_attendance_no.Value = v_attendance_machine_NO;
            OracleParameter[] parameters = new OracleParameter[3] {
                v_num, param_attendance_no,param_year_month_str
            };
            oH.ExecuteNonQuery(proc_Name, parameters);
            return Int32.Parse(parameters[0].Value.ToString());
        }
      
        /// <summary>
        /// 获取某个月，某考勤机的  考勤日期天数。
        /// </summary>
        /// <param name="v_attendance_machine_NO"></param>
        /// <param name="v_year_month_str"></param>
        /// <returns></returns>
        public static int get_nums_of_ar_days(int v_attendance_machine_NO, string v_year_month_str) {
            OracleHelper oH = OracleHelper.getBaseDao();
            string proc_Name = "get_AR_Summary.get_nums_of_ar_days";
            OracleParameter v_num = new OracleParameter("v_num", OracleDbType.Int32, ParameterDirection.ReturnValue);
            OracleParameter param_attendance_no = new OracleParameter("v_attendance_machine_no", OracleDbType.Varchar2, ParameterDirection.Input);
            OracleParameter param_year_month_str = new OracleParameter("v_year_month_str", OracleDbType.Varchar2, ParameterDirection.Input);
            param_year_month_str.Value = v_year_month_str;
            param_attendance_no.Value = v_attendance_machine_NO;
            OracleParameter[] parameters = new OracleParameter[3] {
                v_num, param_attendance_no,param_year_month_str
            };
            oH.ExecuteNonQuery(proc_Name, parameters);
            return Int32.Parse(parameters[0].Value.ToString());
        }
        public static DataTable get_Machines_no(string year_month_str) {
            OracleHelper oH = OracleHelper.getBaseDao();
            string proc_Name = "get_AR_Summary.get_Statff_Info";
            OracleParameter v_cur = new OracleParameter("v_cursor", OracleDbType.RefCursor, ParameterDirection.ReturnValue);
            OracleParameter param_year_month_str = new OracleParameter("v_year_month_str", OracleDbType.Varchar2, ParameterDirection.Input);
            param_year_month_str.Value = year_month_str;
            OracleParameter[] parameters = new OracleParameter[2] {
                v_cur, param_year_month_str
            };
            return oH.getDT(proc_Name, parameters);
        }
        //get_Statff_Info(v_year_month_str varchar2, v_attendance_machine_NO int) return t_cur
        //获得AR_Summary
        /// <summary>
        /// 依据 考勤机号，和年月  获取员工信息。
        /// </summary>
        /// <param name="year_month_str"></param>
        /// <param name="attendance_machine_no"></param>
        /// <returns></returns>
        public static DataTable get_Staff_info(string year_month_str,int attendance_machine_no) {
            OracleHelper oH = OracleHelper.getBaseDao();
            string proc_Name = "get_AR_Summary.get_Statff_Info";
            OracleParameter v_cur = new OracleParameter("v_cursor", OracleDbType.RefCursor, ParameterDirection.ReturnValue);
            OracleParameter param_year_month_str = new OracleParameter("v_year_month_str",OracleDbType.Varchar2,ParameterDirection.Input);
            OracleParameter param_attendance_no = new OracleParameter("v_attendance_machine_no", OracleDbType.Varchar2, ParameterDirection.Input);
            param_year_month_str.Value = year_month_str;
            param_attendance_no.Value = attendance_machine_no;
            OracleParameter[] parameters = new OracleParameter [3]{
                v_cur,param_attendance_no,param_year_month_str 
            };
            return oH.getDT(proc_Name, parameters);
        }
        /// <summary>
        /// 依据年月，考勤机号   获取每人每日的  考勤记录。
        /// </summary>
        /// <param name="year_month_str"></param>
        /// <param name="attendance_machine_no"></param>
        /// <returns></returns>
        public static  DataTable get_AR_Of_Each_Staff(string year_month_str,int attendance_machine_no) {
            OracleHelper oH = OracleHelper.getBaseDao();
            string proc_Name = "get_AR_Summary.get_AR_Of_Each_Staff";
            OracleParameter v_cur = new OracleParameter("v_cursor", OracleDbType.RefCursor, ParameterDirection.ReturnValue);
            OracleParameter param_year_month_str = new OracleParameter("v_year_month_str", OracleDbType.Varchar2, ParameterDirection.Input);
            OracleParameter param_attendance_no = new OracleParameter("v_attendance_machine_no", OracleDbType.Varchar2, ParameterDirection.Input);
            param_year_month_str.Value = year_month_str;
            param_attendance_no.Value = attendance_machine_no;
            OracleParameter[] parameters = new OracleParameter[3]{
                v_cur,param_attendance_no,param_year_month_str 
            };
            return oH.getDT(proc_Name, parameters);
        }
        public static DataTable Get_AR_Summary(string year_month_str, int attendance_machine_no) {
            OracleHelper oH = OracleHelper.getBaseDao();
            string proc_Name = "get_AR_Summary.Get_AR_Summary";
            OracleParameter v_cur = new OracleParameter("v_cursor", OracleDbType.RefCursor, ParameterDirection.ReturnValue);
            OracleParameter param_year_month_str = new OracleParameter("v_year_month_str", OracleDbType.Varchar2, ParameterDirection.Input);
            OracleParameter param_attendance_no = new OracleParameter("v_attendance_machine_no", OracleDbType.Varchar2, ParameterDirection.Input);
            param_year_month_str.Value = year_month_str;
            param_attendance_no.Value = attendance_machine_no;
            OracleParameter[] parameters = new OracleParameter[3]{
                v_cur,param_attendance_no,param_year_month_str 
            };
            return oH.getDT(proc_Name, parameters);
        }
    }
}
