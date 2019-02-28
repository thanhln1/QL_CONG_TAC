using Dapper;
using DTO;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;        
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DAO
{
    public class MT_USERS_DAO
    {
        COMMON dao = new COMMON(); 
        public  List<MT_NHAN_VIEN> LoadUser() {
            using (IDbConnection cnn = new System.Data.SqlClient.SqlConnection(dao.ConnectionString("Default")))
            {
                var output = cnn.Query<MT_NHAN_VIEN>("select * from MT_NHAN_VIEN", new DynamicParameters());
                return output.ToList();
            }
        }

        public  void SaveUser(MT_NHAN_VIEN user) {
            using (IDbConnection cnn = new System.Data.SqlClient.SqlConnection(dao.ConnectionString("Default")))
            {
                cnn.Execute("insert into MT_NHAN_VIEN (MA_NHAN_VIEN, HO_TEN, CHUC_VU, VAI_TRO) values (@MA_NHAN_VIEN, @HO_TEN, @CHUC_VU, @VAI_TRO)", user);   
            }
        }

        public bool checkUserDuplicate( MT_NHAN_VIEN user )
        {
            bool isDuplicate = false;

            using (IDbConnection cnn = new System.Data.SqlClient.SqlConnection(dao.ConnectionString("Default")))
            {
                var output = cnn.Query<MT_HOP_DONG>("select * from MT_NHAN_VIEN a where a.MA_NHAN_VIEN = @MA_NHAN_VIEN ", new { MA_NHAN_VIEN = user.MA_NHAN_VIEN}).ToList();
                if (output.Count > 0)
                {
                    isDuplicate = true;
                }
            }
            return isDuplicate;
        }

        public MT_NHAN_VIEN getLastUser()
        {
            using (IDbConnection cnn = new System.Data.SqlClient.SqlConnection(dao.ConnectionString("Default")))
            {
                var output = cnn.Query<MT_NHAN_VIEN>("select * from MT_NHAN_VIEN a where a.ID = (select Max(ID) from MT_NHAN_VIEN) ", new DynamicParameters());
                return output.First();
            }
        }

        public bool DeleteUser( MT_NHAN_VIEN user )
        {
            try
            {
                using (IDbConnection cnn = new System.Data.SqlClient.SqlConnection(dao.ConnectionString("Default")))
                {
                    StringBuilder sql = new StringBuilder();
                    sql.Append("DELETE FROM MT_NHAN_VIEN ");    
                    sql.Append(" WHERE ID = @ID; "); 
                    cnn.Execute(sql.ToString(), user);
                    return true;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public string getGroupCode( string maNhanVien )
        {
            using (IDbConnection cnn = new System.Data.SqlClient.SqlConnection(dao.ConnectionString("Default")))
            {
                var output = cnn.Query<MT_NHAN_VIEN>("select * from MT_NHAN_VIEN a where a.MA_NHAN_VIEN = @MA_NHAN_VIEN ", new { MA_NHAN_VIEN = maNhanVien });
                if (output.Count() >0)
                {
                    return output.First().PHONG_BAN;
                }
                else
                {
                    return null;
                }

               
            }
        }

        public bool UpdateUser( MT_NHAN_VIEN user )
        {
            try
            {
                using (IDbConnection cnn = new System.Data.SqlClient.SqlConnection(dao.ConnectionString("Default")))
                {
                    StringBuilder sql = new StringBuilder();
                    sql.Append("UPDATE MT_NHAN_VIEN set ");
                    sql.Append("MA_NHAN_VIEN = @MA_NHAN_VIEN, ");
                    sql.Append("HO_TEN = @HO_TEN, ");
                    sql.Append("CHUC_VU = @CHUC_VU, ");
                    sql.Append("VAI_TRO = @VAI_TRO, ");
                    sql.Append("PHONG_BAN = @PHONG_BAN ");
                    sql.Append(" WHERE ID = @ID; ");

                    cnn.Execute(sql.ToString(), user);
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
