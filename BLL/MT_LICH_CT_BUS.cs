using DAO;
using DTO;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BUS
{
   public class MT_LICH_CT_BUS
    {
        MT_LICH_CT_DAO dao = new MT_LICH_CT_DAO();
        public bool SaveCalenda( MT_LICH_CT calenda )
        {
            try
            {
                if (dao.checkCalendaDuplicate(calenda))
                {
                    return false;
                }
                else
                {
                    dao.SaveCalenda(calenda);
                    return true;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            } 
        }

        public MT_LICH_CT getCalenda( int month, int year )
        {
            MT_LICH_CT lichCT = new MT_LICH_CT();
            try
            {
                lichCT = dao.getLichCT(month, year);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return lichCT;
        }
    }
}
