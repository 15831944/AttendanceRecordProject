using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Tools;
namespace AttendanceRecord.Entities
{
    public class AR_Temp
    {
        private string _job_number;
        private string _name;
        private int _attendance_machine_flag;
        private int _row_Index;

       
        public AR_Temp()
        {
        }
        public string Job_number
        {
            get
            {
                return _job_number;
            }

            set
            {
                _job_number = value;
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

        public int Attendance_machine_flag
        {
            get
            {
                return _attendance_machine_flag;
            }

            set
            {
                _attendance_machine_flag = value;
            }
        }

        public int Row_Index
        {
            get
            {
                return _row_Index;
            }

            set
            {
                _row_Index = value;
            }
        }

        //清空临时表
        public static void deleteTheARTemp() {
            string sqlStr = string.Format(@"delete from AR_Temp");
            OracleDaoHelper.executeSQL(sqlStr);
        }
        //保存数据
        public void saveRecord() {
            string sqlStr = string.Format(@"INSERT INTO AR_Temp(ATTENDANCE_MACHINE_FLAG,row_index,job_number,name) values({0},{1},'{2}','{3}')", _attendance_machine_flag,_row_Index, _job_number,_name);
            OracleDaoHelper.executeSQL(sqlStr);
        }
    }
}
