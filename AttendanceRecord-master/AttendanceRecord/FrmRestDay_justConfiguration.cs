using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Tools;
using Oracle.DataAccess.Client;
namespace AttendanceRecord
{
    public partial class FrmRestDay_justConfiguration : Form
    {
        private string _year_and_month;
        private OracleHelper oH = null;

        public FrmRestDay_justConfiguration()
        {
            InitializeComponent();
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

        private void FrmRestDay_justConfiguration_Load(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(_year_and_month))
            {
                return;
            }
            /*
            string procedureName = "PKG_Rest_Day.Generate_Rest_Day";

            if (oH == null)
            {
                oH = OracleHelper.getBaseDao();
            }
            OracleParameter param_Name = new OracleParameter("v_Name", OracleDbType.Varchar2, ParameterDirection.Input);
            OracleParameter param_YearAndMonth = new OracleParameter("v_yearAndMonth", OracleDbType.NVarchar2, ParameterDirection.Input);
            OracleParameter param_restDay_cur = new OracleParameter("v_cur", OracleDbType.RefCursor, ParameterDirection.ReturnValue);
            OracleParameter[] parameters = new OracleParameter[3] { param_restDay_cur, param_Name, param_YearAndMonth };
            parameters[1].Value = "everybody";
            parameters[2].Value = _year_and_month;
            DataTable dt = oH.getDT(procedureName, parameters);
            this.dgv.DataSource = null;
            this.dgv.DataSource = null;
            this.dgv.DataSource = dt;
            DGVHelper.AutoSizeForDGV(dgv);
            */
            //先获取本月最后一天
            int lastDay = 0;
            string sqlStr = string.Format(@"select to_char(last_day(to_date('{0}','yyyy-MM')),'dd') as last_day
                                            from dual",_year_and_month);
            DataTable dt = OracleDaoHelper.getDTBySql(sqlStr);
            lastDay = int.Parse(dt.Rows[0]["last_day"].ToString());

            for (int i = 1; i <= lastDay; i++) {
                string year_and_month_str = _year_and_month + "-" + i.ToString();
                 sqlStr = string.Format(@"SELECT TO_Char(to_date('{0}','yyyy-MM-dd'),'day') the_day
                                            FROM DUAL
                                            ",year_and_month_str);
                dt = OracleDaoHelper.getDTBySql(sqlStr);
                if ("sunday".Equals(dt.Rows[0]["the_day"].ToString().Trim())){
                    sqlStr = string.Format(@"INSERT INTO Rest_Day(name,rest_day,update_time)values('everybody',to_date('{0}','yyyy-MM-dd'),sysdate)",
                                            year_and_month_str);
                    OracleDaoHelper.executeSQL(sqlStr);
                }
            }
            //获取该月得所有休息日
            sqlStr = string.Format(@"SELECT name as ""姓名"",
                                            rest_day as ""休息日"",
                                            update_time as ""更新时间""
                                    FROM Rest_day
                                    WHERE trunc(rest_day,'MM') = to_date('{0}','yyyy-MM')",
                                    _year_and_month);
            dt = OracleDaoHelper.getDTBySql(sqlStr);
            this.dgv.DataSource = dt;
            DGVHelper.AutoSizeForDGV(dgv);
        }

        private void FrmRestDay_justConfiguration_FormClosed(object sender, FormClosedEventArgs e)
        {
            this.DialogResult = DialogResult.OK;
        }
    }
}
