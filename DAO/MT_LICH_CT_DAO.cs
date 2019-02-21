using Dapper;
using DTO;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DAO
{
    public class MT_LICH_CT_DAO
    {
        COMMON dao = new COMMON();
        public void SaveCalenda( MT_LICH_CT calenda )
        {
            using (IDbConnection cnn = new System.Data.SqlClient.SqlConnection(dao.ConnectionString("Default")))
            {
                cnn.Execute("insert into MT_LICH_CT (NAM, THANG, FROM_DATE, TO_DATE) values (@NAM, @THANG, @FROM_DATE, @TO_DATE)", calenda);
            }
        }

        public bool checkCalendaDuplicate( MT_LICH_CT calenda )
        {
            bool isDuplicate = false;

            using (IDbConnection cnn = new System.Data.SqlClient.SqlConnection(dao.ConnectionString("Default")))
            {
                var output = cnn.Query<MT_HOP_DONG>("select * from MT_LICH_CT a where a.THANG = @THANG and a.NAM = @NAM ", new { THANG = calenda.THANG , NAM = calenda.NAM}).ToList();
                if (output.Count > 0)
                {
                    isDuplicate = true;
                }
            }
            return isDuplicate;
        }

        public MT_LICH_CT getLichCT( int month, int year )
        {
            using (IDbConnection cnn = new System.Data.SqlClient.SqlConnection(dao.ConnectionString("Default")))
            {
                var output = cnn.Query<MT_LICH_CT>("select * from MT_LICH_CT a where a.NAM = @YEAR and  a.THANG = @MONTH ", new { YEAR = year, MONTH = month }).ToList();
                if (output.Count > 0)
                {
                    return output.First();
                }
                else
                {
                    return null;
                }
                // return output.First();
            }
        }
    }
}
