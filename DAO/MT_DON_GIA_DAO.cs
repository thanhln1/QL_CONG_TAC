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

        public List<MT_DON_GIA> getAllDonGia()
        {
            using (IDbConnection cnn = new System.Data.SqlClient.SqlConnection(dao.ConnectionString("Default")))
            {
                var output = cnn.Query<MT_DON_GIA>("select * from MT_DON_GIA ").ToList();
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

        public bool UpdateDonGia(MT_DON_GIA dongia)
        {
            try
            {
                using (IDbConnection cnn = new System.Data.SqlClient.SqlConnection(dao.ConnectionString("Default")))
                {
                    StringBuilder sql = new StringBuilder();
                    sql.Append("UPDATE MT_DON_GIA set ");
                    sql.Append("DON_GIA=@DON_GIA, ");
                    sql.Append("GHI_CHU=@GHI_CHU ");
                    sql.Append(" WHERE ID = @ID; ");

                    cnn.Execute(sql.ToString(), dongia);
                    return true;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
   
}
