using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DTO;
using DAO;
using System.Reflection;

namespace BUS
{
    public class MT_SCHEDUAL_BUS
    {
        MT_SCHEDUAL_DAO dao = new MT_SCHEDUAL_DAO();
        MT_LICH_CT_BUS busCalenda = new MT_LICH_CT_BUS();
        public List<VW_SCHEDUAL> loadSchedual( int month, int year, string realOrFake )
        {
            List<VW_SCHEDUAL> listSchedual = new List<VW_SCHEDUAL>();
            try
            {
                listSchedual = dao.LoadSchedual(month, year, realOrFake);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return listSchedual;
        }


        public bool SaveSchedual( MT_SCHEDUAL shedual, int month, int year )
        {
            try
            {
                if (dao.checkSchedualDuplicate(shedual, month, year))
                {
                    return false;
                }
                else
                {
                    dao.SaveSchedual(shedual, month, year);
                    return true;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public List<VW_SCHEDUAL> LoadListSchedual( int month, int year, string realOrFake )
        {
            List<VW_SCHEDUAL> listSchedual = new List<VW_SCHEDUAL>();
            MT_LICH_CT rowCalenda = busCalenda.getCalenda(month, year);
            if (rowCalenda == null)
            {
                return null;
            }
            VW_SCHEDUAL day = new VW_SCHEDUAL();
            day = generateDay(rowCalenda);
            listSchedual.Add(day);

            day = generateThu();
            listSchedual.Add(day);

            List<VW_SCHEDUAL> listNew = new List<VW_SCHEDUAL>();
            listNew = loadSchedual(month, year, realOrFake);
            listSchedual.AddRange(listNew);

            return listSchedual;
        }

        private VW_SCHEDUAL generateDay( MT_LICH_CT rowCalenda )
        {
            DateTime fromDate = rowCalenda.FROM_DATE;
            VW_SCHEDUAL day = new VW_SCHEDUAL();
            day.HO_TEN = "Ngày / Tháng";
            day.ID = 0;
            day.MA_NHAN_VIEN = null;
            day.THANG = 0;
            day.NAM = 0;
            string partem = "dd/MM/yyyy";
            day.TUAN1_THU2 = fromDate.ToString(partem);
            day.TUAN1_THU3 = ( fromDate.AddDays(1) ).ToString(partem);
            day.TUAN1_THU4 = ( fromDate.AddDays(2) ).ToString(partem);
            day.TUAN1_THU5 = ( fromDate.AddDays(3) ).ToString(partem);
            day.TUAN1_THU6 = ( fromDate.AddDays(4) ).ToString(partem);
            day.TUAN1_THU7 = ( fromDate.AddDays(5) ).ToString(partem);
            day.TUAN1_CN = ( fromDate.AddDays(6) ).ToString(partem);
            day.TUAN2_THU2 = ( fromDate.AddDays(7) ).ToString(partem);
            day.TUAN2_THU3 = ( fromDate.AddDays(8) ).ToString(partem);
            day.TUAN2_THU4 = ( fromDate.AddDays(9) ).ToString(partem);
            day.TUAN2_THU5 = ( fromDate.AddDays(10) ).ToString(partem);
            day.TUAN2_THU6 = ( fromDate.AddDays(11) ).ToString(partem);
            day.TUAN2_THU7 = ( fromDate.AddDays(12) ).ToString(partem);
            day.TUAN2_CN = ( fromDate.AddDays(13) ).ToString(partem);
            day.TUAN3_THU2 = ( fromDate.AddDays(14) ).ToString(partem);
            day.TUAN3_THU3 = ( fromDate.AddDays(15) ).ToString(partem);
            day.TUAN3_THU4 = ( fromDate.AddDays(16) ).ToString(partem);
            day.TUAN3_THU5 = ( fromDate.AddDays(17) ).ToString(partem);
            day.TUAN3_THU6 = ( fromDate.AddDays(18) ).ToString(partem);
            day.TUAN3_THU7 = ( fromDate.AddDays(19) ).ToString(partem);
            day.TUAN3_CN = ( fromDate.AddDays(20) ).ToString(partem);
            day.TUAN4_THU2 = ( fromDate.AddDays(21) ).ToString(partem);
            day.TUAN4_THU3 = ( fromDate.AddDays(22) ).ToString(partem);
            day.TUAN4_THU4 = ( fromDate.AddDays(23) ).ToString(partem);
            day.TUAN4_THU5 = ( fromDate.AddDays(24) ).ToString(partem);
            day.TUAN4_THU6 = ( fromDate.AddDays(25) ).ToString(partem);
            day.TUAN4_THU7 = ( fromDate.AddDays(26) ).ToString(partem);
            day.TUAN4_CN = rowCalenda.TO_DATE.ToString(partem);
            return day;
        }

        private VW_SCHEDUAL generateThu()
        {
            VW_SCHEDUAL thu = new VW_SCHEDUAL();
            thu.HO_TEN = "HỌ VÀ TÊN";
            thu.ID = 0;
            thu.MA_NHAN_VIEN = "HỌ VÀ TÊN";
            thu.THANG = 0;
            thu.NAM = 0;
            thu.TUAN1_THU2 = "2";
            thu.TUAN1_THU3 = "3";
            thu.TUAN1_THU4 = "4";
            thu.TUAN1_THU5 = "5";
            thu.TUAN1_THU6 = "6";
            thu.TUAN1_THU7 = "7";
            thu.TUAN1_CN = "CN";
            thu.TUAN2_THU2 = "2";
            thu.TUAN2_THU3 = "3";
            thu.TUAN2_THU4 = "4";
            thu.TUAN2_THU5 = "5";
            thu.TUAN2_THU6 = "6";
            thu.TUAN2_THU7 = "7";
            thu.TUAN2_CN = "CN";
            thu.TUAN3_THU2 = "2";
            thu.TUAN3_THU3 = "3";
            thu.TUAN3_THU4 = "4";
            thu.TUAN3_THU5 = "5";
            thu.TUAN3_THU6 = "6";
            thu.TUAN3_THU7 = "7";
            thu.TUAN3_CN = "CN";
            thu.TUAN4_THU2 = "2";
            thu.TUAN4_THU3 = "3";
            thu.TUAN4_THU4 = "4";
            thu.TUAN4_THU5 = "5";
            thu.TUAN4_THU6 = "6";
            thu.TUAN4_THU7 = "7";
            thu.TUAN4_CN = "CN";
            return thu;
        }

        public List<VW_SCHEDUAL> GetSchedual( int month, int year )
        {
            List<VW_SCHEDUAL> listSchedual = new List<VW_SCHEDUAL>();
            try
            {
                listSchedual = dao.GetSchedual(month, year);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return listSchedual;
        }

        // lấy danh sách nhân viên đi công tác 
        public List<STAFF> GetListStaff( string maKhachHang, int month, int year )
        {
            List<STAFF> listStaffSelect = new List<STAFF>();
            List<VW_SCHEDUAL> listStaff = new List<VW_SCHEDUAL>();
            listStaff = GetSchedual(month, year);

            foreach (VW_SCHEDUAL staff in listStaff)
            {
                List<int> list_ngay_cong_tac = new List<int>();
                STAFF staff_select = new STAFF();
                int count_ngay = 0;

                if (staff.TUAN1_THU2 == maKhachHang) { count_ngay++; list_ngay_cong_tac.Add(0); }
                if (staff.TUAN1_THU3 == maKhachHang) { count_ngay++; list_ngay_cong_tac.Add(1); }
                if (staff.TUAN1_THU4 == maKhachHang) { count_ngay++; list_ngay_cong_tac.Add(2); }
                if (staff.TUAN1_THU5 == maKhachHang) { count_ngay++; list_ngay_cong_tac.Add(3); }
                if (staff.TUAN1_THU6 == maKhachHang) { count_ngay++; list_ngay_cong_tac.Add(4); }
                if (staff.TUAN1_THU7 == maKhachHang) { count_ngay++; list_ngay_cong_tac.Add(5); }
                if (staff.TUAN1_CN == maKhachHang) { count_ngay++; list_ngay_cong_tac.Add(6); }
                if (staff.TUAN2_THU2 == maKhachHang) { count_ngay++; list_ngay_cong_tac.Add(7); }
                if (staff.TUAN2_THU3 == maKhachHang) { count_ngay++; list_ngay_cong_tac.Add(8); }
                if (staff.TUAN2_THU4 == maKhachHang) { count_ngay++; list_ngay_cong_tac.Add(9); }
                if (staff.TUAN2_THU5 == maKhachHang) { count_ngay++; list_ngay_cong_tac.Add(10); }
                if (staff.TUAN2_THU6 == maKhachHang) { count_ngay++; list_ngay_cong_tac.Add(11); }
                if (staff.TUAN2_THU7 == maKhachHang) { count_ngay++; list_ngay_cong_tac.Add(12); }
                if (staff.TUAN2_CN == maKhachHang) { count_ngay++; list_ngay_cong_tac.Add(13); }
                if (staff.TUAN3_THU2 == maKhachHang) { count_ngay++; list_ngay_cong_tac.Add(14); }
                if (staff.TUAN3_THU3 == maKhachHang) { count_ngay++; list_ngay_cong_tac.Add(15); }
                if (staff.TUAN3_THU4 == maKhachHang) { count_ngay++; list_ngay_cong_tac.Add(16); }
                if (staff.TUAN3_THU5 == maKhachHang) { count_ngay++; list_ngay_cong_tac.Add(17); }
                if (staff.TUAN3_THU6 == maKhachHang) { count_ngay++; list_ngay_cong_tac.Add(18); }
                if (staff.TUAN3_THU7 == maKhachHang) { count_ngay++; list_ngay_cong_tac.Add(19); }
                if (staff.TUAN3_CN == maKhachHang) { count_ngay++; list_ngay_cong_tac.Add(20); }
                if (staff.TUAN4_THU2 == maKhachHang) { count_ngay++; list_ngay_cong_tac.Add(21); }
                if (staff.TUAN4_THU3 == maKhachHang) { count_ngay++; list_ngay_cong_tac.Add(22); }
                if (staff.TUAN4_THU4 == maKhachHang) { count_ngay++; list_ngay_cong_tac.Add(23); }
                if (staff.TUAN4_THU5 == maKhachHang) { count_ngay++; list_ngay_cong_tac.Add(24); }
                if (staff.TUAN4_THU6 == maKhachHang) { count_ngay++; list_ngay_cong_tac.Add(25); }
                if (staff.TUAN4_THU7 == maKhachHang) { count_ngay++; list_ngay_cong_tac.Add(26); }
                if (staff.TUAN4_CN == maKhachHang) { count_ngay++; list_ngay_cong_tac.Add(27); }

                if (count_ngay > 0)
                {
                    if (
                         staff.TUAN1_CN == maKhachHang
                       || staff.TUAN1_THU2 == maKhachHang
                       || staff.TUAN1_THU3 == maKhachHang
                       || staff.TUAN1_THU4 == maKhachHang
                       || staff.TUAN1_THU5 == maKhachHang
                       || staff.TUAN1_THU6 == maKhachHang
                       || staff.TUAN1_THU7 == maKhachHang
                       || staff.TUAN1_CN == maKhachHang
                       || staff.TUAN2_THU2 == maKhachHang
                       || staff.TUAN2_THU3 == maKhachHang
                       || staff.TUAN2_THU4 == maKhachHang
                       || staff.TUAN2_THU5 == maKhachHang
                       || staff.TUAN2_THU6 == maKhachHang
                       || staff.TUAN2_THU7 == maKhachHang
                       || staff.TUAN2_CN == maKhachHang
                       || staff.TUAN3_THU2 == maKhachHang
                       || staff.TUAN3_THU3 == maKhachHang
                       || staff.TUAN3_THU4 == maKhachHang
                       || staff.TUAN3_THU5 == maKhachHang
                       || staff.TUAN3_THU6 == maKhachHang
                       || staff.TUAN3_THU7 == maKhachHang
                       || staff.TUAN3_CN == maKhachHang
                       || staff.TUAN4_THU2 == maKhachHang
                       || staff.TUAN4_THU3 == maKhachHang
                       || staff.TUAN4_THU4 == maKhachHang
                       || staff.TUAN4_THU5 == maKhachHang
                       || staff.TUAN4_THU6 == maKhachHang
                       || staff.TUAN4_THU7 == maKhachHang
                       || staff.TUAN4_CN == maKhachHang)
                    {
                        staff_select.HO_TEN = staff.HO_TEN;
                        //staff_select.MA_NHAN_VIEN = staff.MA_NHAN_VIEN;
                        staff_select.SO_NGAY_CONG_TAC = count_ngay;
                        staff_select.NGAY_CONG_TAC = list_ngay_cong_tac;
                        listStaffSelect.Add(staff_select);
                    }
                }
            }
            return listStaffSelect;
        }

        public string[,] GetSchedualArray( DateTime fromDate, DateTime toDate )
        {
            // Tạo 1 mảng 2 chiều với số dòng và số cột đã nhập
            int Rows = 0;
            int Columns = 0;
            string newLine = "\n";
            try
            {
                List<MT_WORKING> listWorking = dao.GetListSchedual(fromDate, toDate);
                if (listWorking.Count() == 0)
                {
                    return null;
                }


                //TimeSpan interval = toDate.Subtract(fromDate);
                //Columns = interval.Days + 3;

                Columns = dao.getColumnFromDate(fromDate, toDate) + 4 ;

                Rows = ( from ld in listWorking select new { id = ld.MA_NHAN_VIEN } ).ToList().Distinct().Count() +2;

                string[,] RsArray = new string[Rows, Columns];

                // Tạo header
                int SoCotDuocThem = 3;
                RsArray[0, 0] = "Họ và Tên";
                RsArray[0, 1] = "Mã Nhân Viên";
                RsArray[0, 2] = "Phòng Ban";
                RsArray[1, 0] = "";
                RsArray[1, 1] = "";
                RsArray[1, 2] = "";
                for (int inx = 0 ; inx < listWorking.Count ; inx++)
                {                      
                    RsArray[0, inx+ SoCotDuocThem] = listWorking[inx].WORKING_DAY.ToString("dd/MM/yyyy");

                    RsArray[1, inx + SoCotDuocThem] = listWorking[inx].WORKING_DAY.DayOfWeek.ToString();
                    // Check đã đến bản ghi nhân viên tiếp theo
                    if (!listWorking[inx].MA_NHAN_VIEN.Equals(listWorking[inx + 1].MA_NHAN_VIEN))
                    {
                        break; 
                    }
                }
                
                // Chỉ số tiếp tục của nhân viên tiếp theo (Bắt đầu từ 2 vì đã thêm 2 dòng tiêu đề bên trên)                         
                int indexContinue = 2;
                // Chỉ số lặp cột của mỗi bản ghi
                int indexColumn = 0;
                //Duyệt List để nhập giá trị cho các phần tử     
                for (int j = 0 ; j < listWorking.Count ; j++)
                {
                    // Kiểm tra nếu là nhân viên khác thì tăng chỉ số
                    if (j>1)
                    {
                        if (!listWorking[j].MA_NHAN_VIEN.Equals(listWorking[j - 1].MA_NHAN_VIEN))
                        {   
                            indexContinue++;
                            indexColumn = 0;
                        }
                    }  

                    // Thêm họ tên, mã nhân viên, phòng ban cho mỗi bản ghi
                    if (indexColumn == 0)
                    {
                        RsArray[indexContinue, 0] = listWorking[j].HO_VA_TEN;
                        RsArray[indexContinue, 1] = listWorking[j].MA_NHAN_VIEN;
                        RsArray[indexContinue, 2] = listWorking[j].PHONG_BAN;
                        string data = listWorking[j].MA_KHACH_HANG+ newLine+ listWorking[j].ID ;
                        RsArray[indexContinue, 3] = data;
                        indexColumn = 3;
                    }
                    else
                    {
                        string data = listWorking[j].MA_KHACH_HANG + newLine + listWorking[j].ID;    
                        RsArray[indexContinue, indexColumn] = data;
                    }
                    indexColumn++;
                }  
                return RsArray; 
            }
            catch (Exception ex)
            { 
                throw ex;
            }
        }              
    }
}
