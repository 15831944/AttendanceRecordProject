using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AttendanceRecord.Entities
{
    public class AR_Sheet_Info:IComparable<AR_Sheet_Info>
    {
        private int _sheetIndex;
        private int _max_rows;

        public AR_Sheet_Info(int _sheetIndex, int _max_rows)
        {
            this._sheetIndex = _sheetIndex;
            this._max_rows = _max_rows;
        }

        public int SheetIndex
        {
            get
            {
                return _sheetIndex;
            }

            set
            {
                _sheetIndex = value;
            }
        }

        public int Max_rows
        {
            get
            {
                return _max_rows;
            }

            set
            {
                _max_rows = value;
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="other"></param>
        /// <returns></returns>
        public int CompareTo(AR_Sheet_Info other)
        {
            if (other == null) return -1;
            if (this.Max_rows > other.Max_rows) return 1;
            return -1;
        }
    }
}
