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

        public bool CheckRunedCalc()
        {
            using (IDbConnection cnn = new System.Data.SqlClient.SqlConnection(dao.ConnectionString("Default")))
            {
                var output = cnn.Query<MT_WORKING>("select * from TMP_WORKING ").ToList();
                if (output.Count > 0)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
        }

        public void SaveSchedual()
        {
            using (IDbConnection cnn = new System.Data.SqlClient.SqlConnection(dao.ConnectionString("Default")))
            {
                // Insert vào bảng HIS, xóa bảng TMP
                cnn.Execute("INSERT INTO HIS_WORKING SELECT * FROM TMP_WORKING ");

                // BackUp bảng MT_HOP_DONG
                cnn.Execute("INSERT INTO HIS_HOP_DONG SELECT * FROM MT_HOP_DONG ");

                // Cập nhật tiền trong hợp đồng.
                cnn.Execute("UPDATE MT_HOP_DONG SET MT_HOP_DONG.CHI_PHI_THUC_DA_CHI = b.CHI_PHI_THUC_DA_CHI FROM MT_HOP_DONG a INNER JOIN TMP_HOP_DONG b ON a.ID = b.ID; ");

                // Xóa Các bảng tạm                     
                cnn.Execute("TRUNCATE TABLE TMP_HOP_DONG;");
                cnn.Execute("TRUNCATE TABLE TMP_WORKING;");
                          
            }
        }

        public void OverWrite()
        {
            using (IDbConnection cnn = new System.Data.SqlClient.SqlConnection(dao.ConnectionString("Default")))
            {
                // Xóa bảng HIS
                cnn.Execute("DELETE FROM HIS_WORKING WHERE ID IN(SELECT ID FROM TMP_WORKING) ");

                // Update bảng HIS từ bảng TMP
                cnn.Execute("INSERT INTO HIS_WORKING SELECT * FROM TMP_WORKING ");
                
                // Xóa Các bảng tạm                     
                cnn.Execute("TRUNCATE TABLE TMP_HOP_DONG;");
                cnn.Execute("TRUNCATE TABLE TMP_WORKING;");

            }
        }

        public bool CheckIsReRun()
        {
            using (IDbConnection cnn = new System.Data.SqlClient.SqlConnection(dao.ConnectionString("Default")))
            {
                var output = cnn.Query<MT_WORKING>("SELECT * FROM HIS_WORKING WHERE ID IN (SELECT ID FROM MT_WORKING) ").ToList();
                if (output.Count > 0)
                {
                    return true;
                }
                else
                {
                    return false;
                }
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
                    sql.Append("MA_KHACH_HANG= @MA_KHACH_HANG, ");
                    sql.Append("MARK= 'FAKE' ");
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
                if (output.Count > 0)
                {
                    return output.First();
                }
                else
                {
                    return null;
                } 
            }
        }
    }
}
