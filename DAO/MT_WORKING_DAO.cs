using System;
using DTO;
using System.Data;
using Dapper;
using System.Linq;
using DAO;
using System.Text;
using System.Collections.Generic;  
using System.Data.SqlClient;
using Dapper;

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
                var output = cnn.Query<MT_WORKING>("select * from TMP_WORKING a where a.ID = @ID ", new { ID = id});
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

        public bool updateWorkingAndContract( MT_WORKING newWorking , string OldMaKH )
        {
            try
            {
                using (IDbConnection cnn = new System.Data.SqlClient.SqlConnection(dao.ConnectionString("Default")))
                {
                    // Check Công ty có còn sử dụng được không?
                    double chiphiconlai = cnn.ExecuteScalar<int>("select (TONG_CHI_PHI_MUC_TOI_DA-CHI_PHI_THUC_DA_CHI) from TMP_HOP_DONG where MA_KHACH_HANG=@MA_KHACH_HANG", new { MA_KHACH_HANG = newWorking.MA_KHACH_HANG });
                    double donGiaNew = cnn.ExecuteScalar<int>("SELECT DON_GIA FROM MT_DON_GIA WHERE DIA_CHI = (SELECT TINH FROM MT_HOP_DONG WHERE MA_KHACH_HANG=@MA_KHACH_HANG);", new { MA_KHACH_HANG = newWorking.MA_KHACH_HANG });
                    double donGiaOld = cnn.ExecuteScalar<int>("SELECT DON_GIA FROM MT_DON_GIA WHERE DIA_CHI = (SELECT TINH FROM MT_HOP_DONG WHERE MA_KHACH_HANG=@MA_KHACH_HANG);", new { MA_KHACH_HANG =OldMaKH });
                    int idNew = cnn.ExecuteScalar<int>("SELECT ID FROM TMP_HOP_DONG WHERE MA_KHACH_HANG = @MA_KHACH_HANG ", new { MA_KHACH_HANG = newWorking.MA_KHACH_HANG });
                    int idOld = cnn.ExecuteScalar<int>("SELECT ID FROM TMP_HOP_DONG WHERE MA_KHACH_HANG = @MA_KHACH_HANG ", new { MA_KHACH_HANG = OldMaKH });

                    if (chiphiconlai >= donGiaNew)
                    {
                        // Update Working
                        StringBuilder sql = new StringBuilder();
                        sql.Append("UPDATE TMP_WORKING set ");
                        sql.Append("MA_KHACH_HANG=@MA_KHACH_HANG ");
                        sql.Append(" WHERE ID = @ID; ");
                        cnn.Execute(sql.ToString(), new { MA_KHACH_HANG = newWorking.MA_KHACH_HANG , ID = newWorking.ID});

                        // Update Contract
                        StringBuilder sql1 = new StringBuilder();
                        // Nếu thay đổi từ null sang có mã KH thì cộng thêm chi phí
                        if (string.IsNullOrWhiteSpace(OldMaKH) && !string.IsNullOrWhiteSpace(newWorking.MA_KHACH_HANG))
                        {
                            sql1.Append("UPDATE TMP_HOP_DONG set ");
                            sql1.Append("CHI_PHI_THUC_DA_CHI=CHI_PHI_THUC_DA_CHI+@DON_GIA ");
                            sql1.Append(" WHERE ID = @ID; ");
                            cnn.Execute(sql1.ToString(), new { DON_GIA = donGiaNew, ID = idNew });
                        }
                        // Nếu thay đổi từ Có mã KH sang null thì trừ chi phí
                        if (!string.IsNullOrWhiteSpace(OldMaKH) && string.IsNullOrWhiteSpace(newWorking.MA_KHACH_HANG))
                        {
                            sql1.Append("UPDATE TMP_HOP_DONG set ");
                            sql1.Append("CHI_PHI_THUC_DA_CHI=CHI_PHI_THUC_DA_CHI-@DON_GIA ");
                            sql1.Append(" WHERE ID =@ID; ");
                            cnn.Execute(sql1.ToString(), new { DON_GIA = donGiaOld, ID = idOld });
                        }
                        // Nếu  chuyển từ mã này sang mã khác thì cập nhật cả hai
                        if (!string.IsNullOrWhiteSpace(OldMaKH) && !string.IsNullOrWhiteSpace(newWorking.MA_KHACH_HANG) && !OldMaKH.Equals(newWorking.MA_KHACH_HANG))
                        {
                            // Trừ tiền công ty cũ
                            sql1.Append("UPDATE TMP_HOP_DONG set ");
                            sql1.Append("CHI_PHI_THUC_DA_CHI=CHI_PHI_THUC_DA_CHI-@DON_GIA_OLD ");
                            sql1.Append(" WHERE ID = @ID_OLD; ");

                            // Cộng tiền công ty mới
                            sql1.Append("UPDATE TMP_HOP_DONG set ");
                            sql1.Append("CHI_PHI_THUC_DA_CHI=CHI_PHI_THUC_DA_CHI+@DON_GIA_NEW ");
                            sql1.Append(" WHERE ID =@ID_NEW; ");

                            cnn.Execute(sql1.ToString(), new { DON_GIA_OLD = donGiaOld, DON_GIA_NEW = donGiaNew, ID_OLD = idOld, ID_NEW= idNew });
                        } 
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                    
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
                var output = cnn.Query<MT_WORKING>("select * from TMP_WORKING order by MA_NHAN_VIEN");
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
                sql.Append("(HO_VA_TEN, MA_NHAN_VIEN, PHONG_BAN, MA_KHACH_HANG, WORKING_DAY, IMPORT_DATE, MARK)");                                        
                sql.Append(" values ");
                sql.Append("(@HO_VA_TEN, @MA_NHAN_VIEN,@PHONG_BAN, @MA_KHACH_HANG, @WORKING_DAY, @IMPORT_DATE, @MARK)");                                                       
                cnn.Execute(sql.ToString(), working);
                return "DONE";

                
            }
        }

        //public string SaveListWorking( List<MT_WORKING> working )
        //{
        //    using (var cnn = new System.Data.SqlClient.SqlConnection(dao.ConnectionString("Default")))
        //    {
        //        cnn.Open();

        //        // create the transaction
        //        // You could use `var` instead of `SqlTransaction`
        //        using (SqlTransaction tran = cnn.BeginTransaction())
        //        {
        //            try
        //            {
        //                foreach (var item in working)
        //                {
        //                    // Check nhóm có đúng không
        //                    var output = cnn.Query<MT_WORKING>("select * from MT_NHAN_VIEN a where a.MA_NHAN_VIEN = @MA_NHAN_VIEN and a.PHONG_BAN = @PHONG_BAN ", new { MA_NHAN_VIEN = working.MA_NHAN_VIEN, PHONG_BAN = working.PHONG_BAN }).ToList();
        //                    if (output.Count < 0)
        //                    {
        //                        // roll the transaction back
        //                        tran.Rollback();

        //                        // handle the error however you need to.
        //                        throw new System.ArgumentException ("NOT_OK");
        //                    } 
        //                    StringBuilder sql = new StringBuilder();
        //                    sql.Append("insert into MT_WORKING ");
        //                    sql.Append("(HO_VA_TEN, MA_NHAN_VIEN, PHONG_BAN, MA_KHACH_HANG, WORKING_DAY, IMPORT_DATE, MARK)");
        //                    sql.Append(" values ");
        //                    sql.Append("(@HO_VA_TEN, @MA_NHAN_VIEN,@PHONG_BAN, @MA_KHACH_HANG, @WORKING_DAY, @IMPORT_DATE, @MARK)");

        //                    // pass the transaction along to the Query, Execute, or the related Async methods. 
        //                    cnn.Execute(sql.ToString(), working,tran);                              
                            
        //                } 
        //                // if it was successful, commit the transaction
        //                tran.Commit();
        //                return "DONE";
        //            }
        //            catch (Exception ex)
        //            {
        //                // roll the transaction back
        //                tran.Rollback();

        //                // handle the error however you need to.
        //                throw ex;
        //            }
        //        }
        //    }

        //}



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

        public List<MT_WORKING> GetWorkingEmpty( DateTime fromCalcDate, DateTime toCalcDate , bool isCN)
        {
            using (IDbConnection cnn = new System.Data.SqlClient.SqlConnection(dao.ConnectionString("Default")))
            {
                // Nếu không chọn ngày chủ nhật thì chạy câu lệnh với điều kiện DATEPART(DW,WORKING_DAY) != '1'
                if (isCN)
                {
                    var output = cnn.Query<MT_WORKING>("select * from TMP_WORKING  where MA_KHACH_HANG='' and cast (WORKING_DAY as date)  between @from and @to and DATEPART(DW,WORKING_DAY) != '1';", new { from = fromCalcDate, to = toCalcDate });
                    return output.ToList();
                }
                // Nếu tính cả CN thì chạy câu lệnh dưới
                else
                {   
                    var output = cnn.Query<MT_WORKING>("select * from TMP_WORKING  where MA_KHACH_HANG='' and cast (WORKING_DAY as date)  between @from and @to;", new { from = fromCalcDate, to = toCalcDate });
                    return output.ToList();
                }  
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