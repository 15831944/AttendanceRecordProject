using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Oracle.DataAccess.Client;
using System.Data;
using Tools;
namespace AttendanceRecord.Entities
{
    public class ARResult
    {
        private string _Year_And_Month_str = String.Empty;
        string procedureName = "Analyze_AR";
        public string Year_And_Month_str
        {
            get
            {
                return _Year_And_Month_str;
            }

            set
            {
                _Year_And_Month_str = value;
            }
        }
        public ARResult(string year_and_month_str) {
            this._Year_And_Month_str = year_and_month_str;
        }
        /// <summary>
        /// 更新分析结果。
        /// </summary>
        /// <returns></returns>
        public int updateARResult(string randomStr) {
            int v_Flag = 0;
            Tools.OracleHelper oH = Tools.OracleHelper.getBaseDao();
            //分析刚刚导入的数据。
            OracleParameter parma_Work_Date = new OracleParameter("v_Work_Date", OracleDbType.NVarchar2, ParameterDirection.Input);
            OracleParameter param_random_Str = new OracleParameter("v_Random_Str", OracleDbType.NVarchar2, ParameterDirection.Input);
            OracleParameter param_Flag = new OracleParameter("v_Flag", OracleDbType.Int32, ParameterDirection.Output);
            parma_Work_Date.Value = this._Year_And_Month_str;
            param_Flag.Value = v_Flag;
            param_random_Str.Value = randomStr;
            OracleParameter[] parameters = new OracleParameter[3] { parma_Work_Date, param_random_Str, param_Flag };
            int j = oH.ExecuteNonQuery(procedureName, parameters);
            return Int32.Parse((parameters[2].Value.ToString()));
        }

    }
}
