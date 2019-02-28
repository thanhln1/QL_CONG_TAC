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
   public class MT_SCHEDUAL_DAO
    {

        COMMON dao = new COMMON();
        public List<VW_SCHEDUAL> LoadSchedual(int month, int year, string realOrFake )
        {
            using (IDbConnection cnn = new System.Data.SqlClient.SqlConnection(dao.ConnectionString("Default")))
            {
                if (realOrFake.Equals("REAL"))
                {
                    var output = cnn.Query<VW_SCHEDUAL>("SELECT B.HO_TEN,A.*   FROM MT_SCHEDUAL as A, MT_NHAN_VIEN as B  Where A.MA_NHAN_VIEN = B.MA_NHAN_VIEN and A.THANG =@MONTH and A.NAM = @YEAR order by A.MA_NHAN_VIEN;", new { MONTH = month, YEAR = year });
                    return output.ToList();
                }
                else if (realOrFake.Equals("FAKE"))
                {
                    var output = cnn.Query<VW_SCHEDUAL>("SELECT B.HO_TEN,A.*   FROM HIS_SCHEDUAL as A, MT_NHAN_VIEN as B  Where A.MA_NHAN_VIEN = B.MA_NHAN_VIEN and A.THANG =@MONTH and A.NAM = @YEAR order by A.MA_NHAN_VIEN;", new { MONTH = month, YEAR = year });
                    return output.ToList();
                }
                else
                {
                    return null;
                }               
               
            }
        }

        public List<MT_SCHEDUAL> LoadSchedual( int month, int year )
        {
            using (IDbConnection cnn = new System.Data.SqlClient.SqlConnection(dao.ConnectionString("Default")))
            {   
                var output = cnn.Query<MT_SCHEDUAL>("SELECT A.*   FROM MT_SCHEDUAL as A  Where A.THANG =@MONTH and A.NAM = @YEAR;", new { MONTH = month, YEAR = year });
                return output.ToList();  
            }
        }

        public bool checkSchedualDuplicate( MT_SCHEDUAL shedual, int month, int year )
        {
            bool isDuplicate = false;

            using (IDbConnection cnn = new System.Data.SqlClient.SqlConnection(dao.ConnectionString("Default")))
            {
                var output = cnn.Query<MT_SCHEDUAL>("select * from MT_SCHEDUAL a where a.MA_NHAN_VIEN = @MA_NHAN_VIEN and a.THANG = @THANG and a.NAM = @NAM", new { MA_NHAN_VIEN = shedual.MA_NHAN_VIEN , THANG = month, NAM = year}).ToList();
                if (output.Count > 0)
                {
                    isDuplicate = true;
                }
            }
            return isDuplicate;
        }


        public void SaveSchedual( MT_SCHEDUAL shedual, int month, int year )
        {
            using (IDbConnection cnn = new System.Data.SqlClient.SqlConnection(dao.ConnectionString("Default")))
            {

                StringBuilder sql = new StringBuilder();
                sql.Append("insert into MT_SCHEDUAL ");
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
    }
}
