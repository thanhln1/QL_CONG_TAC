using System;
using DTO;
using System.Data;
using Dapper;
using System.Linq;
using DAO;
using System.Text;
using System.Collections.Generic;

namespace DAO
{
    public class MT_WORKING_DAO
    {
        COMMON dao = new COMMON();
        public MT_WORKING_DAO()
        {
        }

        public bool checkWorkingDuplicate( MT_WORKING working )
        {
            bool isDuplicate = false;

            using (IDbConnection cnn = new System.Data.SqlClient.SqlConnection(dao.ConnectionString("Default")))
            {
                var output = cnn.Query<MT_WORKING>("select * from MT_WORKING a where a.MA_NHAN_VIEN = @MA_NHAN_VIEN and a.WORKING_DAY = @WORKING_DAY ", new { MA_NHAN_VIEN = working.MA_NHAN_VIEN, WORKING_DAY = working.WORKING_DAY }).ToList();
                if (output.Count > 0)
                {
                    isDuplicate = true;
                }
            }
            return isDuplicate;
        }

        public MT_WORKING GetByID( string id )
        {
            using (IDbConnection cnn = new System.Data.SqlClient.SqlConnection(dao.ConnectionString("Default")))
            {   
                var output = cnn.Query<MT_WORKING>("select * from MT_WORKING a where a.ID = @ID ", new { ID = id});
                if (output!=null || output.Count()>0)
                {
                    return output.First();
                }
                else
                {
                   throw new NullReferenceException("KHONG_CO_DATA");
                }  
            }
        }

        public bool updateWorking( MT_WORKING newWorking )
        {
            try
            {
                using (IDbConnection cnn = new System.Data.SqlClient.SqlConnection(dao.ConnectionString("Default")))
                {
                    StringBuilder sql = new StringBuilder();
                    sql.Append("UPDATE MT_WORKING set ");
                    sql.Append("MA_KHACH_HANG=@MA_KHACH_HANG, "); 
                    sql.Append(" WHERE ID = @ID; ");

                    cnn.Execute(sql.ToString(), newWorking);
                    return true;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public int getColumnFromDateOfREAL( DateTime fromDate, DateTime toDate )
        {
            using (IDbConnection cnn = new System.Data.SqlClient.SqlConnection(dao.ConnectionString("Default")))
            {
                int output = cnn.ExecuteScalar<int>("select  DATEDIFF(day, min(Working_day), max(Working_day) ) from MT_WORKING where cast (WORKING_DAY as date) between @FROM and @TO ", new { FROM = fromDate, TO = toDate });
                return output;
            }
        }

        public int getColumnFromDateofTMP()
        {
            using (IDbConnection cnn = new System.Data.SqlClient.SqlConnection(dao.ConnectionString("Default")))
            {
                int output = cnn.ExecuteScalar<int>("select  DATEDIFF(day, min(Working_day), max(Working_day) ) from TMP_WORKING");
                return output;
            }
        }

        public List<MT_WORKING> GetListTMPSchedual()
        {
            using (IDbConnection cnn = new System.Data.SqlClient.SqlConnection(dao.ConnectionString("Default")))
            {
                var output = cnn.Query<MT_WORKING>("select * from TMP_WORKING order by ID");
                return output.ToList();
            }
        }

        public string SaveWorking( MT_WORKING working )
        {
            using (IDbConnection cnn = new System.Data.SqlClient.SqlConnection(dao.ConnectionString("Default")))
            {
                // Check nhóm có đúng không
                var output = cnn.Query<MT_WORKING>("select * from MT_NHAN_VIEN a where a.MA_NHAN_VIEN = @MA_NHAN_VIEN and a.PHONG_BAN = @PHONG_BAN ", new { MA_NHAN_VIEN = working.MA_NHAN_VIEN, PHONG_BAN = working.PHONG_BAN }).ToList();
                if (output.Count < 0)
                {
                    return "NOT_OK";
                }

                StringBuilder sql = new StringBuilder();
                sql.Append("insert into MT_WORKING ");
                sql.Append("(HO_VA_TEN, MA_NHAN_VIEN, PHONG_BAN, MA_KHACH_HANG, WORKING_DAY, IMPORT_DATE)");                                        
                sql.Append(" values ");
                sql.Append("(@HO_VA_TEN, @MA_NHAN_VIEN,@PHONG_BAN, @MA_KHACH_HANG, @WORKING_DAY, @IMPORT_DATE)");                                                       
                cnn.Execute(sql.ToString(), working);
                return "DONE";

            }
        }

        public void delAllTMP()
        {
            using (IDbConnection cnn = new System.Data.SqlClient.SqlConnection(dao.ConnectionString("Default")))
            {
                cnn.Execute("DELETE FROM TMP_WORKING");
            }
        }
        public void CopySchedual( DateTime fromCalcDate, DateTime toCalcDate )
        {
            using (IDbConnection cnn = new System.Data.SqlClient.SqlConnection(dao.ConnectionString("Default")))
            {
                cnn.Execute("INSERT INTO TMP_WORKING SELECT * FROM MT_WORKING WHERE  cast (WORKING_DAY as date)  between @from and @to;", new { from = fromCalcDate, to = toCalcDate });
            }
        }

        public List<MT_WORKING> GetWorkingEmpty( DateTime fromCalcDate, DateTime toCalcDate )
        {
            using (IDbConnection cnn = new System.Data.SqlClient.SqlConnection(dao.ConnectionString("Default")))
            {
                var output = cnn.Query<MT_WORKING>("select * from TMP_WORKING  where MA_KHACH_HANG='' and cast (WORKING_DAY as date)  between @from and @to;", new { from = fromCalcDate, to = toCalcDate });
                
                return output.ToList();     
            }
        }                                      

        public List<MT_WORKING> GetListRealSchedual( DateTime fromDate, DateTime toDate )
        {
            using (IDbConnection cnn = new System.Data.SqlClient.SqlConnection(dao.ConnectionString("Default")))
            {
                var output = cnn.Query<MT_WORKING>("select * from MT_WORKING where cast (WORKING_DAY as date) between @FROM and @TO order by MA_NHAN_VIEN ", new { FROM = fromDate, TO = toDate });
                return output.ToList();
            }
        }

        public int getColumnFromDateOfFake( DateTime fromDate, DateTime toDate )
        {
            using (IDbConnection cnn = new System.Data.SqlClient.SqlConnection(dao.ConnectionString("Default")))
            {
                int output = cnn.ExecuteScalar<int>("select  DATEDIFF(day, min(Working_day), max(Working_day) ) from HIS_WORKING where cast (WORKING_DAY as date) between @FROM and @TO ", new { FROM = fromDate, TO = toDate });
                return output;
            }
        }

        public List<MT_WORKING> GetListFakeSchedual( DateTime fromDate, DateTime toDate )
        {
            using (IDbConnection cnn = new System.Data.SqlClient.SqlConnection(dao.ConnectionString("Default")))
            {
                var output = cnn.Query<MT_WORKING>("select * from HIS_WORKING where cast (WORKING_DAY as date) between @FROM and @TO order by MA_NHAN_VIEN ", new { FROM = fromDate, TO = toDate });
                return output.ToList();
            }
        }


    }
}