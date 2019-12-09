using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Oracle.DataAccess.Client;
using Tools;
using System.Data;
namespace AttendanceRecord.Entities
{
    public class OverTimeOfRestDay
    {
        //休息日 添加 
        public static int getOverTimeOfRestDay(string ar_jn,string year_and_month) {
            OracleParameter param_OverTime = new OracleParameter("Result", OracleDbType.Int32, ParameterDirection.ReturnValue);
            OracleParameter param_JN = new OracleParameter("v_ar_jn",OracleDbType.Varchar2,ParameterDirection.Input);
            OracleParameter param_Year_And_Month = new OracleParameter("v_year_And_Month",OracleDbType.Varchar2,ParameterDirection.Input);
            param_JN.Size = 50;
            param_Year_And_Month.Size = 20;
            //赋予值.
            param_JN.Value = ar_jn;
            param_Year_And_Month.Value = year_and_month;
            OracleParameter[] parameters = new OracleParameter[3] { param_OverTime, param_JN , param_Year_And_Month  };
            string procedureName = "PKG_AR_DETAIL.GET_OverTime_Of_RestDay";
            OracleHelper oH = OracleHelper.getBaseDao();
            oH.ExecuteNonQuery(procedureName, parameters);
            return int.Parse(parameters[0].Value.ToString());
        }
    }
}
