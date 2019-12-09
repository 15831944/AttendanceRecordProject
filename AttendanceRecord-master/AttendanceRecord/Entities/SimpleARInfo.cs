using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AttendanceRecord.Entities
{
    /// <summary>
    /// 此类 用于 记录 考勤原始数据中，职工姓名，以及所在的考勤机编号。
    /// </summary>
    public class SimpleARInfo
    {
        private  int _attendanceMachineFlag;

        private int _rowIndex;

        private string _name;

        public int RowIndex
        {
            get
            {
                return _rowIndex;
            }

            set
            {
                _rowIndex = value;
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

        public  int AttendanceMachineFlag
        {
            get
            {
                return _attendanceMachineFlag;
            }

            set
            {
                _attendanceMachineFlag = value;
            }
        }

        public string toString() {
            return string.Format(@"{0}号考勤机: 第{1}行,{2}", _attendanceMachineFlag, _rowIndex,_name);
        }
    }
}
