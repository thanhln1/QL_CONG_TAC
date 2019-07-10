using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DTO;
using DAO;

namespace BUS
{
    public class TMP_WORKING_BUS
    {
        TMP_WORKING_DAO dao = new TMP_WORKING_DAO();

        public void DelAllTMP()
        {
            dao.delAllTMP();
        }

        public bool CopySchedual( DateTime fromCalcDate, DateTime toCalcDate )
        {

            try
            {
                dao.CopySchedual(fromCalcDate, toCalcDate);
                return true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

    }
}
