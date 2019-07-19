using System;
using DTO;
using System.Collections.Generic;
using DAO;
using System.Linq;

namespace BUS
{
    public class MT_WORKING_BUS
    {
        MT_WORKING_DAO dao = new MT_WORKING_DAO();
        public static int DAY_OF_WORKING = 3;
        public MT_WORKING_BUS()
        {
        }
        public string SaveWorking( MT_WORKING working )
        {
            try
            {
                if (dao.checkWorkingDuplicate(working))
                {
                    return "DUPLICATE";
                }
                else
                {

                    return dao.SaveWorking(working);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public MT_WORKING GetByID( string id )
        {
            try
            {
                return dao.GetByID(id);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public bool UpdateWorking( MT_WORKING newWorking )
        {
            try
            {
                return dao.updateWorking(newWorking);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public string[,] GetSchedualArray( string RealOrFake, DateTime fromDate, DateTime toDate )
        {
            // Tạo 1 mảng 2 chiều với số dòng và số cột đã nhập
            int Rows = 0;
            int Columns = 0;
            string newLine = "\n";
            try
            {
                List<MT_WORKING> listWorking = new List<MT_WORKING>();
                if (RealOrFake.Equals("REAL"))
                {
                    listWorking = dao.GetListRealSchedual(fromDate, toDate);
                }
                else if (RealOrFake.Equals("FAKE"))
                {
                    listWorking = dao.GetListFakeSchedual(fromDate, toDate);
                }
                else
                {

                }
                if (listWorking.Count == 0)
                {
                    return null;
                }


                //TimeSpan interval = toDate.Subtract(fromDate);
                //Columns = interval.Days + 3;

                Columns = dao.getColumnFromDate(fromDate, toDate) + 4;

                Rows = ( from ld in listWorking select new { id = ld.MA_NHAN_VIEN } ).ToList().Distinct().Count() + 2;

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
                    RsArray[0, inx + SoCotDuocThem] = listWorking[inx].WORKING_DAY.ToString("dd/MM/yyyy");

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
                    if (j > 1)
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
                        string data = listWorking[j].MA_KHACH_HANG + newLine + listWorking[j].ID;
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

        public List<OBJ_CALC> GetWorkingEmpty( DateTime fromCalcDate, DateTime toCalcDate )
        {
            try
            {
                List<OBJ_CALC> ListAvailabe = new List<OBJ_CALC>();

                List<MT_WORKING> listWorking = dao.GetWorkingEmpty(fromCalcDate, toCalcDate);

                string MA_NHAN_VIEN = string.Empty;
                List<FROM_TO> LIST_DAY_NOT_WORKING = new List<FROM_TO>();
                FROM_TO SPACE_DAY = new FROM_TO();
                DateTime FromDate = new DateTime();
                DateTime ToDate = new DateTime();
                // int SO_NGAY_CON_TRONG = 0;

                bool isOtherStaff = false;  
                

                // Bắt đầu duyệt từng phần tử để tạo danh sách còn trống
                for (int i = 0 ; i < listWorking.Count ; i++)
                {
                    // Nếu là phần tử đầu tiên thì set Mã Nhân Viên               
                    if (i == 0)
                    {
                        MA_NHAN_VIEN = listWorking[i].MA_NHAN_VIEN;
                        SPACE_DAY.FromDate = listWorking[i].WORKING_DAY;
                        // LIST_DAY_NOT_WORKING.Add(listWorking[i].WORKING_DAY);
                        // SO_NGAY_CON_TRONG++; 

                        OBJ_CALC item = new OBJ_CALC();
                        item.MA_NHAN_VIEN = listWorking[i].MA_NHAN_VIEN; ;
                        item.LIST_DAY_NOT_WORKING = LIST_DAY_NOT_WORKING;
                        // item.SO_NGAY_CON_TRONG = SO_NGAY_CON_TRONG;
                        ListAvailabe.Add(item);

                    }                                                                            
                    else
                    {
                        // Nếu phần tử tiếp theo vẫn  là nhân viên đó thì cài đặt các thông số
                        if (MA_NHAN_VIEN.Equals(listWorking[i].MA_NHAN_VIEN))
                        {
                            // Kiểm tra tính liên tục giữa 2 ngày
                            TimeSpan diff1 = listWorking[i].WORKING_DAY.Subtract(listWorking[i - 1].WORKING_DAY);
                            // Nếu liên tục thì chỉnh sửa ngày ToDate
                            if (diff1.TotalDays == 1)
                            {                                                                      
                                ToDate = listWorking[i].WORKING_DAY;   
                            }
                            // Nếu không liên tục thì kiểm tra và chèn khoảng ngày làm việc đã tạo vào và thêm khoảng mới
                            else
                            {
                                SPACE_DAY.ToDate = listWorking[i].WORKING_DAY;
                                // Kiểm tra nếu nhiều hơn Số ngày đã cài đặt thì thêm vào danh sách
                                TimeSpan diff2 = SPACE_DAY.ToDate.Subtract(SPACE_DAY.FromDate);
                                // Nếu liên tục thì chỉnh sửa ngày ToDate
                                if (diff1.TotalDays < DAY_OF_WORKING)
                                {
                                    LIST_DAY_NOT_WORKING.Add(SPACE_DAY);
                                }
                                else
                                {
                                    SPACE_DAY.FromDate = listWorking[i].WORKING_DAY;
                                }  
                              
                            }


                            
                        }
                        // Nếu là nhân viên khác thì cài đặt lại các thông số từ đầu
                        else
                        {
                            isOtherStaff = true;
                        }
                        // Nếu không phải nhân viên khác thì tính toán
                        if (!isOtherStaff)
                        {
                            
                        }
                    }
                }

                return ListAvailabe;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}