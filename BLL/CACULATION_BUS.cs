using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using DTO;
using DAO;

namespace BUS
{
    public class CACULATION_BUS
    {                                                     
        MT_CONTRACT_DAO daoContract = new MT_CONTRACT_DAO();
        TMP_WORKING_DAO daoTMP = new TMP_WORKING_DAO();

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

        public string[,] CALC( string[,] SchedualArray, DateTime fromCalcDate, DateTime toCalcDate )
        {
            List<OBJ_CALC> ListCalc = new List<OBJ_CALC>();
            // Duyệt từng row
            for (int i = 2 ; i < SchedualArray.GetLength(0) ; i++)
            {
                OBJ_CALC calc = new OBJ_CALC();
                List<string> listEmpty = new List<string>();
                calc.MA_NHAN_VIEN = SchedualArray[i, 1].ToString();
                // Tạo 1 row là 1 mảng với số cột là Length của phần tử
                string[] row = new string[SchedualArray.GetLength(1)];
                for (int j = 3 ; j < SchedualArray.GetLength(1) ; j++)
                {
                    // bool isNull = false;
                    string data = SchedualArray[i,j].ToString();
                    string id = data.Substring(data.IndexOf('\n') + 1);

                    // Tìm kiếm ngày làm việc để tính toán
                    MT_WORKING oneDay = daoTMP.getByID(id);
                    if (oneDay.MA_KHACH_HANG.Equals("") || string.IsNullOrWhiteSpace(oneDay.MA_KHACH_HANG))
                    {
                        listEmpty.Add(id);
                    }  
                }
               // calc.LIST_DAY_NOT_WORKING = listEmpty;
                ListCalc.Add(calc);                      
             
            }


            // trả về mảng dữ liệu
            throw new NotImplementedException();
        }

        public string[,] CALC( List<OBJ_CALC> DanhSachNgayLamViecConTrong )
        {
            
            throw new NotImplementedException();
        }
    }
}

