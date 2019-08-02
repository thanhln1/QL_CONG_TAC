using System;
using DTO;
using System.Collections.Generic;
using DAO;
using System.Linq;
using System.Diagnostics;

namespace BUS
{
    public class MT_WORKING_BUS
    {
        MT_WORKING_DAO dao = new MT_WORKING_DAO();         
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
                return dao.updateWorkingAndContract(newWorking);
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
                    Columns = dao.getColumnFromDateOfREAL(fromDate, toDate) + 4;
                }
                else if (RealOrFake.Equals("FAKE"))
                {
                    listWorking = dao.GetListFakeSchedual(fromDate, toDate);
                    Columns = dao.getColumnFromDateOfFake(fromDate, toDate) + 4;
                }
                else if (RealOrFake.Equals("TMP"))
                {
                    listWorking = dao.GetListTMPSchedual();
                    Columns = dao.getColumnFromDateofTMP() + 4;
                }  
                if (listWorking.Count == 0)
                {
                    return null;
                }       

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
                        string data = listWorking[j].MA_KHACH_HANG + newLine + listWorking[j].ID + newLine + listWorking[j].MARK;
                        RsArray[indexContinue, 3] = data;
                        indexColumn = 3;
                    }
                    else
                    {
                        string data = listWorking[j].MA_KHACH_HANG + newLine + listWorking[j].ID + newLine + listWorking[j].MARK;
                        RsArray[indexContinue, indexColumn] = data;
                    }
                    indexColumn++;
                }
                return RsArray;
            }
            catch (Exception ex)
            {                     
                // Get stack trace for the exception with source file information
                var st = new StackTrace(ex, true);
                // Get the top stack frame
                var frame = st.GetFrame(0);
                // Get the line number from the stack frame
                var line = frame.GetFileLineNumber();
                throw ex;
            }
        }

        public List<OBJ_CALC> GetWorkingEmpty( DateTime fromCalcDate, DateTime toCalcDate , bool isCN )
        {
            try
            {
                // Danh sách các đối tượng có thể sử dụng
                List<OBJ_CALC> ListAvailabe = new List<OBJ_CALC>();

                // Lấy danh sách các nhân viên có ngày làm việc là trống trong bảng TMP_WORRKING
                List<MT_WORKING> listWorking = dao.GetWorkingEmpty(fromCalcDate, toCalcDate, isCN);

                // Tạo đối tượng để thêm vào danh sách 
                
                string MA_NHAN_VIEN = string.Empty;
                List<List<int>> LIST_DAY_NOT_WORKING = new List<List<int>>();
                List<int> SPACE_DAY = new List<int>();   
                DateTime OldDate = new DateTime();   

                // Bắt đầu duyệt từng phần tử để tạo danh sách còn trống
                for (int i = 0 ; i < listWorking.Count ; i++)
                {
                    
                    // Nếu là phần tử đầu tiên thì set Mã Nhân Viên               
                    if (i == 0)
                    {
                        MA_NHAN_VIEN = listWorking[i].MA_NHAN_VIEN;
                        OldDate = listWorking[i].WORKING_DAY;
                        // Thêm Id vào danh sách đã tạo
                        SPACE_DAY.Add(listWorking[i].ID);
                    }                                                                          
                    else
                    {                          
                        // Nếu phần tử tiếp theo vẫn  là nhân viên đó thì cài đặt các thông số
                        if (MA_NHAN_VIEN.Equals(listWorking[i].MA_NHAN_VIEN))
                        {
                            // Kiểm tra tính liên tục giữa 2 ngày
                            TimeSpan diff1 = listWorking[i].WORKING_DAY.Subtract(listWorking[i - 1].WORKING_DAY);
                            // Nếu liên tục thì thêm Id vào danh sách đã tạo
                            if (diff1.TotalDays == 1)
                            {                                
                                SPACE_DAY.Add(listWorking[i].ID);
                            }
                            // Nếu không liên tục thì kiểm tra và chèn khoảng ngày làm việc đã tạo vào và thêm khoảng mới
                            else
                            {                                  
                                // Kiểm tra nếu nhiều hơn Số ngày đã cài đặt thì thêm vào danh sách  
                                if (SPACE_DAY.Count >= COMMON_BUS.DAY_OF_WORKING)
                                {
                                    List<int> listID = SPACE_DAY.ToList();
                                    LIST_DAY_NOT_WORKING.Add(listID);
                                }
                                SPACE_DAY.Clear();
                                SPACE_DAY.Add(listWorking[i].ID);
                                                               
                            }
                            OldDate = listWorking[i].WORKING_DAY;
                        }
                        // Nếu là nhân viên khác thì tạo đối tượng và thêm vào danh sách
                        else
                        {
                            if (LIST_DAY_NOT_WORKING.Count >0)
                            {
                                // Tạo đối tượng mới và chèn vào danh sách  
                                OBJ_CALC newObject = new OBJ_CALC();
                                newObject.MA_NHAN_VIEN = MA_NHAN_VIEN;
                                List<int> listID = SPACE_DAY.ToList();
                                LIST_DAY_NOT_WORKING.Add(listID);
                                newObject.LIST_DAY_NOT_WORKING = LIST_DAY_NOT_WORKING.ToList();
                                ListAvailabe.Add(newObject);

                                LIST_DAY_NOT_WORKING.Clear();
                            }                      
                            

                            // Sau khi chèn danh sách thì tạo mới đối tượng cho nhân viên tiếp theo
                            MA_NHAN_VIEN = listWorking[i].MA_NHAN_VIEN;
                            SPACE_DAY.Clear();
                            LIST_DAY_NOT_WORKING.Clear();
                            OldDate = listWorking[i].WORKING_DAY;
                            // Thêm Id vào danh sách đã tạo
                            SPACE_DAY.Add(listWorking[i].ID);

                        }   
                    }
                    if (i== (listWorking.Count-1))
                    {                                         
                        // Kiểm tra nếu nhiều hơn Số ngày đã cài đặt thì thêm vào danh sách  
                        if (SPACE_DAY.Count >= COMMON_BUS.DAY_OF_WORKING)
                        {
                            List<int> listID = SPACE_DAY.ToList();
                            LIST_DAY_NOT_WORKING.Add(listID);

                            OBJ_CALC newObject = new OBJ_CALC();
                            newObject.MA_NHAN_VIEN = MA_NHAN_VIEN;
                            newObject.LIST_DAY_NOT_WORKING = LIST_DAY_NOT_WORKING.ToList();
                            ListAvailabe.Add(newObject);
                                                            
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