using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using Tools;
namespace AttendanceRecord.Entities
{
    public class MESSAGES
    {
        public static DataTable getMSG(string randomStr) {
            string sqlStr = String.Format(@"
                                            select 
                                                   prompt, 
                                                    flag, 
                                                    ipaddr, 
                                                    operate_time, 
                                                    subject, 
                                                    random_str 
                                            from Message
                                            WHERE random_Str = '{0}'", randomStr);
            return OracleDaoHelper.getDTBySql(sqlStr);
        } 
    }
}
