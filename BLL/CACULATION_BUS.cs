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
            // Duyệt từng row
            for (int i = 0 ; i < SchedualArray.GetLength(0) ; i++)
            {

                // Tạo 1 row là 1 mảng với số cột là Length của phần tử
                string[] row = new string[SchedualArray.GetLength(1)];
                for (int j = 0 ; j < SchedualArray.GetLength(1) ; j++)
                {
                    bool isNull = false;
                    string data = SchedualArray[i,j].ToString();
                    string id = data.Substring(data.IndexOf('\n') + 1);

                    // Tìm kiếm ngày làm việc để tính toán
                    MT_WORKING oneDay = daoTMP.getByID(id);
                    if (oneDay.MA_KHACH_HANG.Equals("") || string.IsNullOrWhiteSpace(oneDay.MA_KHACH_HANG))
                    {
                        if (!isNull)
                        {
                            isNull = true;
                        }
                    
                    }
                    else
                    {
                        if (isNull)
                        {
                            isNull = false;
                        }

                    }
                }                           
             
            }

            // trả về mảng dữ liệu
            throw new NotImplementedException();
        }
    }
}

