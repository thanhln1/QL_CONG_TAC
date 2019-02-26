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
    public class MT_DON_GIA_DAO
    {
        COMMON dao = new COMMON();
        public List<MT_DON_GIA> getDonGia(string diachi)
        {
            using (IDbConnection cnn = new System.Data.SqlClient.SqlConnection(dao.ConnectionString("Default")))
            {
                var output = cnn.Query<MT_DON_GIA>("select * from MT_DON_GIA a where a.DIA_CHI = @DIA_CHI ", new { DIA_CHI = diachi }).ToList();
                if (output.Count > 0)
                {
                    return output.ToList();
                }
                else
                {
                    return null;
                }
            }
        }
    }
   
}
