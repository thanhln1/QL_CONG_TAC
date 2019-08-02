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
   public class TMP_WORKING_DAO
    {

        COMMON dao = new COMMON();

        public void delAllTMP()
        {
            using (IDbConnection cnn = new System.Data.SqlClient.SqlConnection(dao.ConnectionString("Default")))
            {
                cnn.Execute("TRUNCATE TABLE TMP_WORKING");
                cnn.Execute("TRUNCATE TABLE TMP_HOP_DONG");
            }
        }
        public void CopySchedual( DateTime fromCalcDate, DateTime toCalcDate )
        {
            using (IDbConnection cnn = new System.Data.SqlClient.SqlConnection(dao.ConnectionString("Default")))
            {
                cnn.Execute("INSERT INTO TMP_WORKING SELECT * FROM MT_WORKING WHERE  cast (WORKING_DAY as date) between @from and @to;", new { from = fromCalcDate, to = toCalcDate });
            }
        }

        public void BackUpHD()
        {
            using (IDbConnection cnn = new System.Data.SqlClient.SqlConnection(dao.ConnectionString("Default")))
            {
                cnn.Execute("INSERT INTO TMP_HOP_DONG SELECT * FROM MT_HOP_DONG;");
            }
        }

        public MT_WORKING getByID( string id )
        {
            using (IDbConnection cnn = new System.Data.SqlClient.SqlConnection(dao.ConnectionString("Default")))
            {
                var output = cnn.Query<MT_WORKING>("select * from MT_WORKING  where ID=@ID", new { ID = id });
                return output.First();
            }
        }

        public DateTime getDayByID( int id1 )
        {
            using (IDbConnection cnn = new System.Data.SqlClient.SqlConnection(dao.ConnectionString("Default")))
            {
                var output = cnn.Query<MT_WORKING>("select * from TMP_WORKING a  where a.ID = @ID ", new { ID = id1 }).ToList();
                
                return output.First().WORKING_DAY;  
            }
        }

        public void UpdateChiPhi( int ID, double ChiPhiPhatSinh )
        {
            try
            {
                using (IDbConnection cnn = new System.Data.SqlClient.SqlConnection(dao.ConnectionString("Default")))
                {
                    StringBuilder sql = new StringBuilder();
                    sql.Append("UPDATE TMP_HOP_DONG set ");
                    sql.Append("CHI_PHI_THUC_DA_CHI=CHI_PHI_THUC_DA_CHI + @CHI_PHI_PHAT_SINH ");  
                    sql.Append(" WHERE ID = @ID; ");

                    cnn.Execute(sql.ToString(), new { ID = ID , CHI_PHI_PHAT_SINH = ChiPhiPhatSinh });  
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void UpdateCompanyToID( int ID, string mA_KHACH_HANG )
        {
            try
            {
                using (IDbConnection cnn = new System.Data.SqlClient.SqlConnection(dao.ConnectionString("Default")))
                {
                    StringBuilder sql = new StringBuilder();
                    sql.Append("UPDATE TMP_WORKING set ");
                    sql.Append("MA_KHACH_HANG= @MA_KHACH_HANG ");
                    sql.Append(" WHERE ID = @ID; ");

                    cnn.Execute(sql.ToString(), new { ID = ID, MA_KHACH_HANG = mA_KHACH_HANG });
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public MT_NHAN_VIEN GetUserByIdOfTMP( string id )
        {
            using (IDbConnection cnn = new System.Data.SqlClient.SqlConnection(dao.ConnectionString("Default")))
            {
                var output = cnn.Query<MT_NHAN_VIEN>("SELECT * from MT_NHAN_VIEN where MA_NHAN_VIEN = (select MA_NHAN_VIEN from TMP_WORKING where ID = @ID)", new { ID = id }).ToList();

                return output.First();

            }
        }
    }
}
