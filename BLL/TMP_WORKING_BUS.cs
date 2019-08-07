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

        public void BackupHD()
        {
            try
            {
                dao.BackUpHD();  
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public bool CheckRunedCalc()
        {
            try
            {
                bool hasRecord = dao.CheckRunedCalc(); 
                return hasRecord;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public bool saveCalc()
        {
            try
            {
                dao.SaveSchedual();
                return true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public bool CheckReRunCALC()
        {
            try
            {
                bool IsReRun = dao.CheckIsReRun();
                return IsReRun;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public bool OverWrite()
        {
            try
            {
                dao.OverWrite();
                return true;
            }
            catch (Exception ex)
            {                  
                throw ex;
            }
        }
    }
}
