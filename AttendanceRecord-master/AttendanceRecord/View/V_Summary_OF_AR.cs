using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using Oracle.DataAccess.Client;
using Tools;
namespace AttendanceRecord
{
    public class V_Summary_OF_AR
    {
        private string _jN;
        private string _year_and_month;
        private string _name;
        public V_Summary_OF_AR(string _jN, string _year_and_month)
        {
            this.JN = _jN;
            this.Year_and_month = _year_and_month;
        }
        /// <summary>
        /// 用于解决 job_number中含有"_"的情形。
        /// </summary>
        /// <returns></returns>
        public string getJN() {
            int index = _jN.IndexOf("_");
            if (index >= 0) {
                return _jN.Substring(index + 1);
            }
            return _jN;
        }
        public string JN
        {
            get
            {
                return _jN;
            }

            set
            {
                _jN = value;
            }
        }

        public string Year_and_month
        {
            get
            {
                return _year_and_month;
            }

            set
            {
                _year_and_month = value;
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
        /// <summary>
        /// PKG_Get_Summary_Of_AR.get_summary_of_ar
        /// </summary>
        /// <returns></returns>
        public DataTable getSummaryOFAR() {
            string procedureName = "PKG_Get_Summary_Of_AR.get_summary_of_ar";
            OracleHelper oH = OracleHelper.getBaseDao();    
          
            //分析刚刚导入的数据。
            OracleParameter param_JN = new OracleParameter("v_job_number", OracleDbType.Varchar2, ParameterDirection.Input);
            OracleParameter param_Name = new OracleParameter("v_name", OracleDbType.Varchar2, ParameterDirection.Input);
            OracleParameter param_year_and_Month = new OracleParameter("v_year_and_month", OracleDbType.Varchar2, ParameterDirection.Input);
            OracleParameter param_cur_result = new OracleParameter("v_cur_result", OracleDbType.RefCursor, ParameterDirection.ReturnValue);
            //parma_JN.Value = this.getJN();
            param_JN.Value = this._jN;
            param_Name.Value = this._name;
            param_year_and_Month.Value = this._year_and_month;
            
            OracleParameter[] parameters = new OracleParameter[4] { param_cur_result, param_JN,param_Name, param_year_and_Month };
            DataTable dt = oH.getDT(procedureName, parameters);
            return dt;
        }
        /// <summary>
        /// PKG_Get_Summary_Of_AR.get_summary_of_ar
        /// </summary>
        /// <returns></returns>
        public static void generateAttendanceRecordBriefly(string year_and_month_str)
        {
            //summarize_attendance_record(v_year_and_month);   
            string procedureName = "summarize_attendance_record";
            OracleHelper oH = OracleHelper.getBaseDao();
            OracleParameter param_year_and_Month = new OracleParameter("v_year_and_month", OracleDbType.Varchar2, ParameterDirection.Input);
            //parma_JN.Value = this.getJN();
            param_year_and_Month.Value = year_and_month_str;
            OracleParameter[] parameters = new OracleParameter[1] { param_year_and_Month };
            oH.ExecuteNonQuery(procedureName, parameters);
        }
        public DataTable getSummaryOfAttendance_Record_Summary()
        {
            string procedureName = "pkg_to_export_from_a_r_summary.get_summary_of_A_R_Summary";
            OracleHelper oH = OracleHelper.getBaseDao();

            //分析刚刚导入的数据。
            OracleParameter parma_JN = new OracleParameter("v_job_number", OracleDbType.Varchar2, ParameterDirection.Input);
            OracleParameter param_year_and_Month = new OracleParameter("v_year_and_month", OracleDbType.Varchar2, ParameterDirection.Input);
            OracleParameter param_cur_result = new OracleParameter("v_cur_result", OracleDbType.RefCursor, ParameterDirection.ReturnValue);
            parma_JN.Value = this._jN;
            param_year_and_Month.Value = this._year_and_month;

            OracleParameter[] parameters = new OracleParameter[3] { param_cur_result, parma_JN, param_year_and_Month };
            DataTable dt = oH.getDT(procedureName, parameters);
            return dt;
        }
    }
}
