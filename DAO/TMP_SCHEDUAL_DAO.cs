﻿using Dapper;
using DTO;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DAO
{
   public class TMP_SCHEDUAL_DAO
    {

        COMMON dao = new COMMON();
        public List<VW_SCHEDUAL> LoadSchedual()
        {
            using (IDbConnection cnn = new System.Data.SqlClient.SqlConnection(dao.ConnectionString("Default")))
            {                  
                var output = cnn.Query<VW_SCHEDUAL>("SELECT B.HO_TEN,A.*   FROM TMP_SCHEDUAL as A, MT_NHAN_VIEN as B  Where A.MA_NHAN_VIEN = B.MA_NHAN_VIEN order by A.MA_NHAN_VIEN;", new DynamicParameters());
                return output.ToList();   
            }
        }

        public List<TMP_SCHEDUAL> LoadSchedual1( int month, int year )
        {
            using (IDbConnection cnn = new System.Data.SqlClient.SqlConnection(dao.ConnectionString("Default")))
            {   
                var output = cnn.Query<TMP_SCHEDUAL>("SELECT A.*   FROM TMP_SCHEDUAL as A  Where A.THANG =@MONTH and A.NAM = @YEAR;", new { MONTH = month, YEAR = year });
                return output.ToList();  
            }
        }

        public bool checkSchedualDuplicate( MT_SCHEDUAL shedual, int month, int year )
        {
            bool isDuplicate = false;

            using (IDbConnection cnn = new System.Data.SqlClient.SqlConnection(dao.ConnectionString("Default")))
            {
                var output = cnn.Query<MT_SCHEDUAL>("select * from TMP_SCHEDUAL a where a.MA_NHAN_VIEN = @MA_NHAN_VIEN and a.THANG = @THANG and a.NAM = @NAM", new { MA_NHAN_VIEN = shedual.MA_NHAN_VIEN , THANG = month, NAM = year}).ToList();
                if (output.Count > 0)
                {
                    isDuplicate = true;
                }
            }
            return isDuplicate;
        }


        public void SaveSchedual( MT_SCHEDUAL shedual )
        {
            using (IDbConnection cnn = new System.Data.SqlClient.SqlConnection(dao.ConnectionString("Default")))
            {

                StringBuilder sql = new StringBuilder();
                sql.Append("insert into TMP_SCHEDUAL ");
                sql.Append("(MA_NHAN_VIEN, THANG, NAM,");
                sql.Append(" TUAN1_THU2, TUAN1_THU3, TUAN1_THU4, TUAN1_THU5, TUAN1_THU6, TUAN1_THU7, TUAN1_CN,");
                sql.Append(" TUAN2_THU2, TUAN2_THU3, TUAN2_THU4, TUAN2_THU5, TUAN2_THU6, TUAN2_THU7, TUAN2_CN,");
                sql.Append(" TUAN3_THU2, TUAN3_THU3, TUAN3_THU4, TUAN3_THU5, TUAN3_THU6, TUAN3_THU7, TUAN3_CN,");
                sql.Append(" TUAN4_THU2, TUAN4_THU3, TUAN4_THU4, TUAN4_THU5, TUAN4_THU6, TUAN4_THU7, TUAN4_CN) ");
                sql.Append(" values "); 
                sql.Append("(@MA_NHAN_VIEN, @THANG, @NAM,");
                sql.Append(" @TUAN1_THU2, @TUAN1_THU3, @TUAN1_THU4, @TUAN1_THU5, @TUAN1_THU6, @TUAN1_THU7, @TUAN1_CN,");
                sql.Append(" @TUAN2_THU2, @TUAN2_THU3, @TUAN2_THU4, @TUAN2_THU5, @TUAN2_THU6, @TUAN2_THU7, @TUAN2_CN,");
                sql.Append(" @TUAN3_THU2, @TUAN3_THU3, @TUAN3_THU4, @TUAN3_THU5, @TUAN3_THU6, @TUAN3_THU7, @TUAN3_CN,");
                sql.Append(" @TUAN4_THU2, @TUAN4_THU3, @TUAN4_THU4, @TUAN4_THU5, @TUAN4_THU6, @TUAN4_THU7, @TUAN4_CN) ");
                cnn.Execute(sql.ToString(), shedual);
            }
        }

        public List<VW_SCHEDUAL> GetSchedual(int month, int year)
        {
            using (IDbConnection cnn = new System.Data.SqlClient.SqlConnection(dao.ConnectionString("Default")))
            {
                // var output = cnn.Query<VW_SCHEDUAL>("SELECT * FROM HIS_SCHEDUAL A Where A.THANG = @THANG and A.NAM = @NAM ;", new {THANG = month, NAM = year});
                var output = cnn.Query<VW_SCHEDUAL>("select NHANVIEN.HO_TEN, SCHEDUAL.* from HIS_SCHEDUAL SCHEDUAL INNER JOIN MT_NHAN_VIEN NHANVIEN ON SCHEDUAL.MA_NHAN_VIEN = NHANVIEN.MA_NHAN_VIEN where SCHEDUAL.THANG = @THANG and SCHEDUAL.NAM = @NAM", new { THANG = month, NAM = year });
                return output.ToList();
            }
        }

        public void delAllTMP()
        {
            using (IDbConnection cnn = new System.Data.SqlClient.SqlConnection(dao.ConnectionString("Default")))
            {                  
                cnn.Execute("DELETE FROM TMP_WORKING");
            }
        }

        public bool checkExitCalc( int month, int year )
        {
            using (IDbConnection cnn = new System.Data.SqlClient.SqlConnection(dao.ConnectionString("Default")))
            {
                var output = cnn.Query<MT_SCHEDUAL>("SELECT *   FROM HIS_SCHEDUAL as s  Where s.THANG = @THANG and s.NAM = @NAM", new { THANG = month, NAM = year });
                return ( output.Count() > 0) ? true : false;               
            }
        }

        public void OverWiteCalc(int month, int year)
        {
            using (IDbConnection cnn = new System.Data.SqlClient.SqlConnection(dao.ConnectionString("Default")))
            {
                // xem lại lỗi chỗ này
                cnn.Execute("DELETE HIS_SCHEDUAL Where THANG = @THANG and NAM = @NAM", new { THANG = month, NAM = year });

               cnn.Execute("INSERT INTO HIS_SCHEDUAL SELECT * FROM TMP_SCHEDUAL;");
            }
        }

        public void saveCalc()
        {
            using (IDbConnection cnn = new System.Data.SqlClient.SqlConnection(dao.ConnectionString("Default")))
            {     
                cnn.Execute("INSERT INTO HIS_SCHEDUAL SELECT * FROM TMP_SCHEDUAL;");
            }
        }

        public bool OverWiteContract( List<MT_HOP_DONG> listTmpHopDong )
        {
            try
            {
                using (IDbConnection cnn = new System.Data.SqlClient.SqlConnection(dao.ConnectionString("Default")))
                {
                    foreach (var contract in listTmpHopDong)
                    {
                        StringBuilder sql = new StringBuilder();
                        sql.Append("UPDATE MT_HOP_DONG set ");
                        //sql.Append("SO_HOP_DONG=@SO_HOP_DONG, ");
                        //sql.Append("NGAY_HOP_DONG=@NGAY_HOP_DONG, ");
                        //sql.Append("NGAY_THANH_LY=@NGAY_THANH_LY, ");
                        //sql.Append("KHACH_HANG=@KHACH_HANG, ");
                        //sql.Append("MA_KHACH_HANG=@MA_KHACH_HANG, ");
                        //sql.Append("NHOM_KHACH_HANG=@NHOM_KHACH_HANG, ");
                        //sql.Append("DIA_CHI=@DIA_CHI, ");
                        //sql.Append("TINH=@TINH, ");
                        //sql.Append("GIA_TRI_HOP_DONG=@GIA_TRI_HOP_DONG, ");
                        //sql.Append("TONG_CHI_PHI_MUC_TOI_DA=@TONG_CHI_PHI_MUC_TOI_DA, ");
                        sql.Append("CHI_PHI_THUC_DA_CHI=@CHI_PHI_THUC_DA_CHI ");
                        // sql.Append("GHI_CHU=@GHI_CHU ");
                        sql.Append(" WHERE ID = @ID AND SO_HOP_DONG=@SO_HOP_DONG; ");

                        cnn.Execute(sql.ToString(), contract); 
                    }
                    
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
