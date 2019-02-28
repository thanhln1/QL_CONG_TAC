using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DTO;
using DAO;

namespace BUS
{
    public class TMP_SCHEDUAL_BUS
    {
        TMP_SCHEDUAL_DAO dao = new TMP_SCHEDUAL_DAO();
        MT_LICH_CT_BUS busCalenda = new MT_LICH_CT_BUS();
        public List<VW_SCHEDUAL> loadSchedual()
        {
            List<VW_SCHEDUAL> listSchedual = new List<VW_SCHEDUAL>();
            try
            {
                listSchedual = dao.LoadSchedual();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return listSchedual;
        }


        public bool SaveSchedual( MT_SCHEDUAL shedual, int month, int year)
        {
            try
            {
                if (dao.checkSchedualDuplicate(shedual, month, year))
                {
                    return false;
                }
                else
                {
                    dao.SaveSchedual(shedual);
                    return true;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public List<VW_SCHEDUAL> LoadListSchedual( int month, int year)
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
            listNew = loadSchedual();
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

        public List<VW_SCHEDUAL> GetSchedual(int month, int year)
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

        public void DelAllTMP()
        {
            dao.delAllTMP();
        }
    }
}
