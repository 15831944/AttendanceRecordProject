using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using Oracle.DataAccess.Client;
namespace AttendanceRecord.Entities
{
    public class WorkSummary
    {
        private string _workDate;
        string procedureName = "Generate_Work_Summary";
        public string WorkDate
        {
            get
            {
                return _workDate;
            }

            set
            {
                _workDate = value;
            }
        }
        public WorkSummary(string dateStr) {
            this._workDate = dateStr;
        }
        //获取相应的工作安排记录
        public DataTable getARSummary() {
            DataTable dt = null;
            string sqlStr = String.Format(@"
                                              SELECt 
                                                seq, 
                                                work_and_rest_date,                                                
                                                dept,
                                                on_duty_time,
                                                off_duty_time,
                                                work_or_rest,
                                                fp_number,
                                                record_time
                                          FROM Work_Summary
                                          WHERE TRUNC(work_And_Rest_Date,'MM') = TO_DATE('{0}','yyyy-MM')
                                         OrDER BY work_and_rest_date DESC,
                                                dept desc", this.WorkDate);
            dt = Tools.OracleDaoHelper.getDTBySql(sqlStr);
            this.convertColName(dt);
            return dt;
        }
        /// <summary>
        /// 生成工作安排。
        /// </summary>
        public int generateARSummary() {
            Tools.OracleHelper oH = Tools.OracleHelper.getBaseDao();
            //分析刚刚导入的数据。
            OracleParameter parma_Work_Date = new OracleParameter("v_Work_Date", OracleDbType.NVarchar2, ParameterDirection.Input);
            parma_Work_Date.Value = this._workDate;
            OracleParameter[] parameters = new OracleParameter[1] { parma_Work_Date };
            int j = oH.ExecuteNonQuery(procedureName, parameters);
            return j;
        }
        public  void convertColName(DataTable dt) {
            dt.Columns["SEQ"].ColumnName = Properties.Resources.SEQ;
            dt.Columns["Dept"].ColumnName = Properties.Resources.DEPT;
            dt.Columns["ON_DUTY_TIME"].ColumnName = Properties.Resources.ON_DUTY_TIME;
            dt.Columns["RECORD_TIME"].ColumnName = Properties.Resources.RECORD_TIME;
            dt.Columns["WORK_AND_REST_DATE"].ColumnName = Properties.Resources.WORK_AND_REST_DATE; 
            dt.Columns["WORK_OR_REST"].ColumnName = Properties.Resources.WORK_OR_REST;
            dt.Columns["OFF_DUTY_TIME"].ColumnName = Properties.Resources.OFF_DUTY_TIME;
            dt.Columns["FP_NUMBER"].ColumnName = Properties.Resources.FP_NUMBER; 
        }

        #region 判断是否有相应的考勤记录
        /// <summary>
        /// 判断是否有相应的考勤记录
        /// </summary>
        /// <param name="yearAndMonthStr"></param>
        /// <returns></returns>
        public static bool ifExistsWorkSummary (string yearAndMonthStr){
            string sqlStr = String.Format(@"  SELECT 
                                                1
                                            FROM Work_Summary
                                            WHERE TRUNC(WORK_AND_REST_DATE,'MM') = TO_DATE('{0}','YYYY-MM')", yearAndMonthStr);
            return Tools.OracleDaoHelper.getDTBySql(sqlStr).Rows.Count > 0 ? true : false;
         }
        #endregion
    }
}
