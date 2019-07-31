﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using DTO;
using DAO;
using System.Text.RegularExpressions;

namespace BUS
{
    public class CACULATION_BUS
    {
        MT_CONTRACT_DAO daoContract = new MT_CONTRACT_DAO();
        TMP_WORKING_DAO daoTMP = new TMP_WORKING_DAO();
        MT_DON_GIA_BUS busDonGia = new MT_DON_GIA_BUS();
        MT_WORKING_BUS busWorking = new MT_WORKING_BUS();

        public List<MT_HOP_DONG> getListCompanyNotFinished()
        {
            List<MT_HOP_DONG> listContract = new List<MT_HOP_DONG>();
            try
            {
                listContract = daoContract.getListCompanyNotFinished();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return listContract;
        }

        //public string[,] CALC( string[,] SchedualArray, DateTime fromCalcDate, DateTime toCalcDate )
        //{
        //    List<OBJ_CALC> ListCalc = new List<OBJ_CALC>();
        //    // Duyệt từng row
        //    for (int i = 2 ; i < SchedualArray.GetLength(0) ; i++)
        //    {
        //        OBJ_CALC calc = new OBJ_CALC();
        //        List<string> listEmpty = new List<string>();
        //        calc.MA_NHAN_VIEN = SchedualArray[i, 1].ToString();
        //        // Tạo 1 row là 1 mảng với số cột là Length của phần tử
        //        string[] row = new string[SchedualArray.GetLength(1)];
        //        for (int j = 3 ; j < SchedualArray.GetLength(1) ; j++)
        //        {
        //            // bool isNull = false;
        //            string data = SchedualArray[i,j].ToString();
        //            string id = data.Substring(data.IndexOf('\n') + 1);

        //            // Tìm kiếm ngày làm việc để tính toán
        //            MT_WORKING oneDay = daoTMP.getByID(id);
        //            if (oneDay.MA_KHACH_HANG.Equals("") || string.IsNullOrWhiteSpace(oneDay.MA_KHACH_HANG))
        //            {
        //                listEmpty.Add(id);
        //            }  
        //        }
        //       // calc.LIST_DAY_NOT_WORKING = listEmpty;
        //        ListCalc.Add(calc);                      

        //    }


        //    // trả về mảng dữ liệu
        //    throw new NotImplementedException();
        //}

        public List<List<string>> CALC( List<OBJ_CALC> DanhSachNgayLamViecConTrong )
        {
            List<List<string>> ListDayMatch = new List<List<string>>();
            try
            {
                List<OBJ_CALC> ListCalc = new List<OBJ_CALC>(DanhSachNgayLamViecConTrong);

                // Chạy lần lượt danh sách
                foreach (var item in DanhSachNgayLamViecConTrong)
                {   // Xóa phần tử đầu tiên
                    ListCalc.Remove(item);
                    // Chạy lần lượt từng danh sách ngày trống trong danh sách
                    foreach (var Compare1 in item.LIST_DAY_NOT_WORKING)
                    {
                        // Chạy lần lượt từng item trong danh sách đã xóa phần tử đầu tiên
                        foreach (var item2 in ListCalc)
                        {
                            // Chạy lần lượt từng danh sách ngày trống trong danh sách đã xóa phần tử đầu tiên
                            foreach (var Compare2 in item2.LIST_DAY_NOT_WORKING)
                            {
                                // Kiểm tra hai danh sách làm việc có trùng nhau không. Nếu trùng thì thêm vào danh sách hai ngày làm việc giống nhau 
                                List<string> DayMatch = ComparesWorkDay(Compare1, Compare2);
                                if (DayMatch.Count >= COMMON_BUS.DAY_OF_WORKING)
                                {
                                    ListDayMatch.Add(DayMatch);
                                }

                            }

                        }

                    }
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }
            return ListDayMatch;

        }

        public List<string> ComparesWorkDay( List<int> listInput1, List<int> listInput2 )
        {
            List<string> dayMatch = new List<string>();
            foreach (var Id1 in listInput1)
            {
                string day = Id1.ToString();
                DateTime day1 = daoTMP.getDayByID(Id1);
                foreach (var Id2 in listInput2)
                {
                    DateTime day2 = daoTMP.getDayByID(Id2);
                    if (DateTime.Compare(day1, day2) == 0)
                    {
                        day += ";" + Id2;
                    }
                }
                if (day.Contains(";"))
                {
                    dayMatch.Add(day);
                }
                else
                {
                    day = string.Empty;
                }

            }
            return dayMatch;
        }
        /// <summary>
        /// Set Company and return result
        /// </summary>
        /// <param name="fakeSchedualArray"></param>
        /// <param name="listCompany"></param>
        /// <returns></returns>
        public string[,] SetCompany( List<List<string>> fakeSchedualArray )
        {
            try
            {
                foreach (var item in fakeSchedualArray)
                {
                    // Số hàng cần set giống nhau
                    int a = Regex.Split(item.First(), ";").Length;
                    // Tính số ngày trống
                    int numDayNull = a * item.Count;                     

                    // Lấy danh sách các công ty chưa hết kinh phí từ bảng tạm
                    List<MT_HOP_DONG> listCompany = getListCompanyNotFinished();                    
                    foreach (var company in listCompany)
                    {
                        // Chạy danh sách các công ty có thể sử dụng 
                        double dongia = busDonGia.GetDonGia(company.TINH);
                        double giaTriCan = dongia * numDayNull;
                        double giaTriKhaDung = company.GIA_TRI_HOP_DONG - company.TONG_CHI_PHI_MUC_TOI_DA;

                        // Nếu giá trị khả dụng lớn hơn giá trị cần.
                        if (giaTriKhaDung >= giaTriCan)
                        {
                            foreach (var groupID in item)
                            {
                                // Cắt chuỗi lấy ra danh sách các ID để Update công ty                                                
                                string[] arrID = Regex.Split(groupID, ";");   
                            
                                // Set công ty vào các chỗ trống trong bảng tạm
                                foreach (var id in arrID)
                                {
                                    daoTMP.UpdateCompanyToID(Int32.Parse(id), company.MA_KHACH_HANG);
                                }                                  
                            }
                            // Update bảng chi phí
                            daoTMP.UpdateChiPhi(company.ID, giaTriCan);
                            break;
                        } 
                    }
                }
                // Lấy danh sách sau khi đã set ra màn hình
                string[,] tmpSchedualArray = busWorking.GetSchedualArray("TMP", DateTime.Now, DateTime.Now);
                return tmpSchedualArray;

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}

