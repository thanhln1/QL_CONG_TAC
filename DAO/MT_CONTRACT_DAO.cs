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
   public class MT_CONTRACT_DAO
    {
        COMMON dao = new COMMON();
        public List<MT_HOP_DONG> LoadContract()
        {
            using (IDbConnection cnn = new System.Data.SqlClient.SqlConnection(dao.ConnectionString("Default")))
            {
                var output = cnn.Query<MT_HOP_DONG>("select * from MT_HOP_DONG", new DynamicParameters());
                return output.ToList();
            }
        }

        public void SaveContract( MT_HOP_DONG contract )
        {
            using (IDbConnection cnn = new System.Data.SqlClient.SqlConnection(dao.ConnectionString("Default")))
            {

                StringBuilder sql = new StringBuilder();
                sql.Append("insert into MT_HOP_DONG ");
                sql.Append("(SO_HOP_DONG,  NGAY_HOP_DONG ,  NGAY_THANH_LY,  KHACH_HANG  , MA_KHACH_HANG,  NHOM_KHACH_HANG,  DIA_CHI,  TINH,  GIA_TRI_HOP_DONG,  TONG_CHI_PHI_MUC_TOI_DA,  CHI_PHI_THUC_DA_CHI,  GHI_CHU) ");
                sql.Append(" values ");
                sql.Append("(@SO_HOP_DONG, @NGAY_HOP_DONG , @NGAY_THANH_LY, @KHACH_HANG , @MA_KHACH_HANG, @NHOM_KHACH_HANG, @DIA_CHI, @TINH, @GIA_TRI_HOP_DONG, @TONG_CHI_PHI_MUC_TOI_DA, @CHI_PHI_THUC_DA_CHI, @GHI_CHU )");
                cnn.Execute(sql.ToString(), contract);
            }
        }

        public List<MT_HOP_DONG> getListCompanyNotFinished()
        {
            using (IDbConnection cnn = new System.Data.SqlClient.SqlConnection(dao.ConnectionString("Default")))
            {
                var output = cnn.Query<MT_HOP_DONG>("select * from TMP_HOP_DONG where CHI_PHI_THUC_DA_CHI < TONG_CHI_PHI_MUC_TOI_DA", new DynamicParameters());
                return output.ToList();
            }
        }

        public bool DeleteContract( MT_HOP_DONG contract )
        {
            try
            {
                using (IDbConnection cnn = new System.Data.SqlClient.SqlConnection(dao.ConnectionString("Default")))
                {
                    StringBuilder sql = new StringBuilder();
                    sql.Append("DELETE FROM MT_HOP_DONG ");
                    sql.Append(" WHERE ID = @ID; ");
                    cnn.Execute(sql.ToString(), contract);
                    return true;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public string getGroupCode( string maKhachHang )
        {
            using (IDbConnection cnn = new System.Data.SqlClient.SqlConnection(dao.ConnectionString("Default")))
            {
                var output = cnn.Query<MT_HOP_DONG>("select * from MT_HOP_DONG a where a.MA_KHACH_HANG = @MA_KHACH_HANG ", new { MA_KHACH_HANG = maKhachHang });
                return output.First().NHOM_KHACH_HANG;
            }
        }

        public bool UpdateContract( MT_HOP_DONG contract )
        {
            try
            {
                using (IDbConnection cnn = new System.Data.SqlClient.SqlConnection(dao.ConnectionString("Default")))
                {
                    StringBuilder sql = new StringBuilder();
                    sql.Append("UPDATE MT_HOP_DONG set "); 
                    sql.Append("SO_HOP_DONG=@SO_HOP_DONG, ");
                    sql.Append("NGAY_HOP_DONG=@NGAY_HOP_DONG, ");
                    sql.Append("NGAY_THANH_LY=@NGAY_THANH_LY, ");
                    sql.Append("KHACH_HANG=@KHACH_HANG, ");
                    sql.Append("MA_KHACH_HANG=@MA_KHACH_HANG, ");
                    sql.Append("NHOM_KHACH_HANG=@NHOM_KHACH_HANG, ");
                    sql.Append("DIA_CHI=@DIA_CHI, ");
                    sql.Append("TINH=@TINH, ");
                    sql.Append("GIA_TRI_HOP_DONG=@GIA_TRI_HOP_DONG, ");
                    sql.Append("TONG_CHI_PHI_MUC_TOI_DA=@TONG_CHI_PHI_MUC_TOI_DA, ");
                    sql.Append("CHI_PHI_THUC_DA_CHI=@CHI_PHI_THUC_DA_CHI, ");
                    sql.Append("GHI_CHU=@GHI_CHU ");
                    sql.Append(" WHERE ID = @ID; ");

                    cnn.Execute(sql.ToString(), contract);
                    return true;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public bool checkContractDuplicate(MT_HOP_DONG contract)
        {
            bool isDuplicate = false;

            using (IDbConnection cnn = new System.Data.SqlClient.SqlConnection(dao.ConnectionString("Default")))
            {
                var output = cnn.Query<MT_HOP_DONG>("select * from MT_HOP_DONG a where a.SO_HOP_DONG = @SO_HOP_DONG ", new { @SO_HOP_DONG = contract.SO_HOP_DONG }).ToList();
                if (output.Count > 0)
                {
                    isDuplicate = true;
                }
            }
            return isDuplicate;
        }

        public MT_HOP_DONG getLastContract()
        {
            using (IDbConnection cnn = new System.Data.SqlClient.SqlConnection(dao.ConnectionString("Default")))
            {
                var output = cnn.Query<MT_HOP_DONG>("select * from MT_HOP_DONG a where a.ID = (select Max(ID) from MT_HOP_DONG) ", new DynamicParameters());
                return output.First();
            }
        }

        // xuất quyết định, bảng kê - danh sách nhân viên - Thanh
        public List<MT_HOP_DONG> GetInforContract(string hopdong)
        {
            using (IDbConnection cnn = new System.Data.SqlClient.SqlConnection(dao.ConnectionString("Default")))
            {
                var output = cnn.Query<MT_HOP_DONG>("select * from MT_HOP_DONG a where a.MA_KHACH_HANG = @MA_KHACH_HANG", new { @MA_KHACH_HANG = hopdong });
                return output.ToList();
            }
        }
    }
}
