using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Oracle.DataAccess.Client;
using System.Data;
using Tools;
namespace AttendanceRecord.Entities
{
    public class ARSummaryFinal
    {
        /// <summary>
        /// 获取最终的 考勤汇总：所属部门	所属组	实际出勤天数 请假时长	未打卡	平日延点	加班日工作时长	加班合计(小时)	迟到	早退	餐补	备注
        /// </summary>
        /// <param name="v_year_and_month_str"></param>
        /// <returns></returns>
        public static DataTable getARSummaryFinal(string v_year_and_month_str ) {
            string procName = "PKG_AR_SUMMARY.getARSummary";
            OracleParameter param_year_and_month = new OracleParameter("v_year_and_month_str",OracleDbType.Varchar2,ParameterDirection.Input);
            OracleParameter param_cur_result = new OracleParameter("v_cur_result", OracleDbType.RefCursor, ParameterDirection.ReturnValue);
            param_year_and_month.Value = v_year_and_month_str;
            OracleParameter[] parameters = new OracleParameter[2] { param_cur_result,param_year_and_month  };
            OracleHelper oH = OracleHelper.getBaseDao();
            return oH.getDT(procName, parameters); 
       }
    }
}
