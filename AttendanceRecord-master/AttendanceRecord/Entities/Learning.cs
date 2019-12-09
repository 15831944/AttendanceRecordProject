using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using Tools;
namespace AttendanceRecord.Entities
{
    public class Learning
    {
        private string _name;

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
        public Learning(string name) {
            this._name = name;
        }

        public bool ifExists() {
            string sqlStr = string.Format(@"SELECT 1 FROM LEARNING WHERE NAME = '{0}'",this._name);
            return OracleDaoHelper.getDTBySql(sqlStr).Rows.Count > 0 ? true : false;
        }

        public int add() {
            string sqlStr = string.Format(@"INSERT INTO LEARNING (NAME) VALUES ('{0}')", this._name);
            return OracleDaoHelper.executeSQL(sqlStr);
        }

        public int del() {
            string sqlStr = string.Format(@"DELETE FROM  LEARNING WHERE NAME = '{0}'", this._name);
            return OracleDaoHelper.executeSQL(sqlStr);
        }

        public static DataTable getAllLearners() {
            string sqlStr = string.Format(@"SELECT NAME AS ""姓名"" FROM LEARNING");
            return OracleDaoHelper.getDTBySql(sqlStr);
        }
    }
}
