using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DTO;
using DAO;

namespace BUS
{
    public class MT_SCHEDUAL_BUS
    {
        MT_SCHEDUAL_DAO dao = new MT_SCHEDUAL_DAO();
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
                    dao.SaveSchedual(shedual, month, year);
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
