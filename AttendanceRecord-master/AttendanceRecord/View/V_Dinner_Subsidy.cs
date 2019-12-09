using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Oracle.DataAccess.Client;
using System.Data;
using Tools;
namespace AttendanceRecord.View
{
    public class V_Dinner_Subsidy
    {
        public static int getDinnerSubsidyNum(string _jN, string _yearAndMonth)
        {
            string procedureName = "get_dinner_subsidy_num";
            OracleParameter param_JN = new OracleParameter("v_JOB_NUMBER", OracleDbType.Varchar2, ParameterDirection.Input);
            OracleParameter param_year_and_month = new OracleParameter("v_year_and_month", OracleDbType.Varchar2, ParameterDirection.Input);
            OracleParameter param_vacation_hours = new OracleParameter("v_Dinner_Subsidy_Num", OracleDbType.Int32, ParameterDirection.Output);
            OracleParameter[] parameters = new OracleParameter[3] { param_JN, param_year_and_month, param_vacation_hours };
            parameters[0].Value = _jN;
            parameters[1].Value = _yearAndMonth;
            OracleHelper oH = OracleHelper.getBaseDao();
            oH.ExecuteNonQuery(procedureName, parameters);
            return int.Parse(parameters[2].Value.ToString());
        }
    }
}
