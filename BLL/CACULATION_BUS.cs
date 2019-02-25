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
        MT_SCHEDUAL_DAO daoSchedual = new MT_SCHEDUAL_DAO();
        MT_CONTRACT_DAO daoContract = new MT_CONTRACT_DAO();
        public List<MT_SCHEDUAL> getListSchedual( int month, int year )
        {
            List<MT_SCHEDUAL> listSchedual = new List<MT_SCHEDUAL>();
            try
            {
                listSchedual = daoSchedual.LoadSchedual(month, year);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return listSchedual;
        }

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
    }
}

