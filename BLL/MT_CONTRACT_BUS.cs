using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DTO;
using DAO;

namespace BUS
{
    public class MT_CONTRACT_BUS
    {
        MT_CONTRACT_DAO dao = new MT_CONTRACT_DAO();
        public bool SaveContract( MT_HOP_DONG contract )
        {
            try
            {
                if (dao.checkContractDuplicate(contract))
                {
                    return false;
                }
                else
                {
                    dao.SaveContract(contract);
                    return true;
                }
            }
            catch (Exception ex)
            {   
                throw ex;
            }              
        }

        public List<MT_HOP_DONG> GetListContract()
        {
            List<MT_HOP_DONG> listContract = new List<MT_HOP_DONG>();
            try
            {
                listContract = dao.LoadContract();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return listContract;
        }

        public bool DelContract( MT_HOP_DONG contract )
        {
            bool isDeleted = false;
            try
            {
                isDeleted = dao.DeleteContract(contract);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return isDeleted;
        }

        public bool UpdateContract( MT_HOP_DONG contract )
        {
            bool isUpdate = false;
            try
            {
                isUpdate = dao.UpdateContract(contract);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return isUpdate;
        }

        // xuất quyết định, bảng kê - danh sách nhân viên - Thanh
        public MT_HOP_DONG GetInforContract(string maKhachHang)
        {                                              
            try
            {
                return dao.GetInforContract(maKhachHang);
            }
            catch (Exception ex)
            {
                throw ex;
            }                 
        }

        public string getGroupCompany( string maKhachHang )
        {
            string groupCode;
            try
            {
                groupCode = dao.getGroupCode(maKhachHang);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return groupCode;
        }

        public MT_HOP_DONG GetInforContractByMaHD( string maKhachHang )
        {
            MT_HOP_DONG contract = new MT_HOP_DONG();
            try
            {
                contract = dao.GetInforContractByMaHD(maKhachHang);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return contract;
        }

        public List<MT_HOP_DONG> checkListContractDuplicate( List<MT_HOP_DONG> listContract )
        {
            List<MT_HOP_DONG> listContractDuplicate = new List<MT_HOP_DONG>();
            foreach (var contract in listContract)
            {  
                if (dao.checkContractDuplicate(contract))
                {
                    listContractDuplicate.Add(contract);
                }                
            }
            return listContractDuplicate;
        }
        public int SaveListContract(List<MT_HOP_DONG> listContract)
        {
            try
            {   
                return dao.SaveListContract(listContract);
              
            }
            catch (Exception ex)
            {
                throw ex;
            }              
        }
    }
}
