using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DTO;
using DAO;

namespace BUS
{
    public class HIS_SCHEDUAL_BUS
    {
        HIS_SCHEDUAL_DAO dao = new HIS_SCHEDUAL_DAO();
        public bool SaveHisSchedual(HIS_SCHEDUAL shedual, int month, int year)
        {
            try
            {           
                dao.SaveSchedual(shedual, month, year);
                return true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
