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

        public object GetListContract()
        {
            List<MT_HOP_DONG> listUser = new List<MT_HOP_DONG>();
            try
            {
                listUser = dao.LoadContract();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return listUser;
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
        public List<MT_HOP_DONG> GetInforContract(string maKhachHang)
        {
            List<MT_HOP_DONG> listUser = new List<MT_HOP_DONG>();
            try
            {
                listUser = dao.GetInforContract(maKhachHang);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return listUser;
        }
    }
}
