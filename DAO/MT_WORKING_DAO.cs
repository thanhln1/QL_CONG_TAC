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
                var output = cnn.Query<MT_WORKING>("select * from MT_WORKING  where MA_KHACH_HANG='' and cast (WORKING_DAY as date)  between @from and @to;", new { from = fromCalcDate, to = toCalcDate });
                
                return output.ToList();     
            }
        }

        //public List<STAFF> GetListStaff( string maKhachHang, int month, int year )
        //{
        //    List<STAFF> listStaffSelect = new List<STAFF>();
        //    List<VW_SCHEDUAL> listStaff = new List<VW_SCHEDUAL>();
        //    listStaff = GetSchedual(month, year);

        //    foreach (VW_SCHEDUAL staff in listStaff)
        //    {
        //        List<int> list_ngay_cong_tac = new List<int>();
        //        STAFF staff_select = new STAFF();
        //        int count_ngay = 0;

        //        if (staff.TUAN1_THU2 == maKhachHang) { count_ngay++; list_ngay_cong_tac.Add(0); }
        //        if (staff.TUAN1_THU3 == maKhachHang) { count_ngay++; list_ngay_cong_tac.Add(1); }
        //        if (staff.TUAN1_THU4 == maKhachHang) { count_ngay++; list_ngay_cong_tac.Add(2); }
        //        if (staff.TUAN1_THU5 == maKhachHang) { count_ngay++; list_ngay_cong_tac.Add(3); }
        //        if (staff.TUAN1_THU6 == maKhachHang) { count_ngay++; list_ngay_cong_tac.Add(4); }
        //        if (staff.TUAN1_THU7 == maKhachHang) { count_ngay++; list_ngay_cong_tac.Add(5); }
        //        if (staff.TUAN1_CN == maKhachHang) { count_ngay++; list_ngay_cong_tac.Add(6); }
        //        if (staff.TUAN2_THU2 == maKhachHang) { count_ngay++; list_ngay_cong_tac.Add(7); }
        //        if (staff.TUAN2_THU3 == maKhachHang) { count_ngay++; list_ngay_cong_tac.Add(8); }
        //        if (staff.TUAN2_THU4 == maKhachHang) { count_ngay++; list_ngay_cong_tac.Add(9); }
        //        if (staff.TUAN2_THU5 == maKhachHang) { count_ngay++; list_ngay_cong_tac.Add(10); }
        //        if (staff.TUAN2_THU6 == maKhachHang) { count_ngay++; list_ngay_cong_tac.Add(11); }
        //        if (staff.TUAN2_THU7 == maKhachHang) { count_ngay++; list_ngay_cong_tac.Add(12); }
        //        if (staff.TUAN2_CN == maKhachHang) { count_ngay++; list_ngay_cong_tac.Add(13); }
        //        if (staff.TUAN3_THU2 == maKhachHang) { count_ngay++; list_ngay_cong_tac.Add(14); }
        //        if (staff.TUAN3_THU3 == maKhachHang) { count_ngay++; list_ngay_cong_tac.Add(15); }
        //        if (staff.TUAN3_THU4 == maKhachHang) { count_ngay++; list_ngay_cong_tac.Add(16); }
        //        if (staff.TUAN3_THU5 == maKhachHang) { count_ngay++; list_ngay_cong_tac.Add(17); }
        //        if (staff.TUAN3_THU6 == maKhachHang) { count_ngay++; list_ngay_cong_tac.Add(18); }
        //        if (staff.TUAN3_THU7 == maKhachHang) { count_ngay++; list_ngay_cong_tac.Add(19); }
        //        if (staff.TUAN3_CN == maKhachHang) { count_ngay++; list_ngay_cong_tac.Add(20); }
        //        if (staff.TUAN4_THU2 == maKhachHang) { count_ngay++; list_ngay_cong_tac.Add(21); }
        //        if (staff.TUAN4_THU3 == maKhachHang) { count_ngay++; list_ngay_cong_tac.Add(22); }
        //        if (staff.TUAN4_THU4 == maKhachHang) { count_ngay++; list_ngay_cong_tac.Add(23); }
        //        if (staff.TUAN4_THU5 == maKhachHang) { count_ngay++; list_ngay_cong_tac.Add(24); }
        //        if (staff.TUAN4_THU6 == maKhachHang) { count_ngay++; list_ngay_cong_tac.Add(25); }
        //        if (staff.TUAN4_THU7 == maKhachHang) { count_ngay++; list_ngay_cong_tac.Add(26); }
        //        if (staff.TUAN4_CN == maKhachHang) { count_ngay++; list_ngay_cong_tac.Add(27); }

        //        if (count_ngay > 0)
        //        {
        //            if (
        //                 staff.TUAN1_CN == maKhachHang
        //               || staff.TUAN1_THU2 == maKhachHang
        //               || staff.TUAN1_THU3 == maKhachHang
        //               || staff.TUAN1_THU4 == maKhachHang
        //               || staff.TUAN1_THU5 == maKhachHang
        //               || staff.TUAN1_THU6 == maKhachHang
        //               || staff.TUAN1_THU7 == maKhachHang
        //               || staff.TUAN1_CN == maKhachHang
        //               || staff.TUAN2_THU2 == maKhachHang
        //               || staff.TUAN2_THU3 == maKhachHang
        //               || staff.TUAN2_THU4 == maKhachHang
        //               || staff.TUAN2_THU5 == maKhachHang
        //               || staff.TUAN2_THU6 == maKhachHang
        //               || staff.TUAN2_THU7 == maKhachHang
        //               || staff.TUAN2_CN == maKhachHang
        //               || staff.TUAN3_THU2 == maKhachHang
        //               || staff.TUAN3_THU3 == maKhachHang
        //               || staff.TUAN3_THU4 == maKhachHang
        //               || staff.TUAN3_THU5 == maKhachHang
        //               || staff.TUAN3_THU6 == maKhachHang
        //               || staff.TUAN3_THU7 == maKhachHang
        //               || staff.TUAN3_CN == maKhachHang
        //               || staff.TUAN4_THU2 == maKhachHang
        //               || staff.TUAN4_THU3 == maKhachHang
        //               || staff.TUAN4_THU4 == maKhachHang
        //               || staff.TUAN4_THU5 == maKhachHang
        //               || staff.TUAN4_THU6 == maKhachHang
        //               || staff.TUAN4_THU7 == maKhachHang
        //               || staff.TUAN4_CN == maKhachHang)
        //            {
        //                staff_select.HO_TEN = staff.HO_TEN;
        //                //staff_select.MA_NHAN_VIEN = staff.MA_NHAN_VIEN;
        //                staff_select.SO_NGAY_CONG_TAC = count_ngay;
        //                staff_select.NGAY_CONG_TAC = list_ngay_cong_tac;
        //                listStaffSelect.Add(staff_select);
        //            }
        //        }
        //    }
        //    return listStaffSelect;
        //}



        public List<MT_WORKING> GetListRealSchedual( DateTime fromDate, DateTime toDate )
        {
            using (IDbConnection cnn = new System.Data.SqlClient.SqlConnection(dao.ConnectionString("Default")))
            {
                var output = cnn.Query<MT_WORKING>("select * from MT_WORKING where cast (WORKING_DAY as date) between @FROM and @TO order by MA_NHAN_VIEN ", new { FROM = fromDate, TO = toDate });
                return output.ToList();
            }
        }

        public int getColumnFromDate( DateTime fromDate, DateTime toDate )
        {
            using (IDbConnection cnn = new System.Data.SqlClient.SqlConnection(dao.ConnectionString("Default")))
            {
                int output = cnn.ExecuteScalar<int>("select  DATEDIFF(day, min(Working_day), max(Working_day) ) from MT_WORKING where cast (WORKING_DAY as date) between @FROM and @TO ", new { FROM = fromDate, TO = toDate });
                return output;
            }
        }

        public List<MT_WORKING> GetListFakeSchedual( DateTime fromDate, DateTime toDate )
        {
            using (IDbConnection cnn = new System.Data.SqlClient.SqlConnection(dao.ConnectionString("Default")))
            {
                var output = cnn.Query<MT_WORKING>("select * from TMP_WORKING where cast (WORKING_DAY as date) between @FROM and @TO order by MA_NHAN_VIEN ", new { FROM = fromDate, TO = toDate });
                return output.ToList();
            }
        }


    }
}