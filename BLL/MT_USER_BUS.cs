using DAO;
using DTO;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BUS
{
    public class MT_USER_BUS
    {
        MT_USERS_DAO dao = new MT_USERS_DAO();
        public List<MT_NHAN_VIEN> GetListUser()
        {
            List<MT_NHAN_VIEN> listUser = new List<MT_NHAN_VIEN>();
            try
            {
                listUser = dao.LoadUser(); 
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return listUser;
        }

        public bool SaveUser( MT_NHAN_VIEN user )
        {                                           
            try
            {
                if (dao.checkUserDuplicate(user))
                {
                    return false;
                }
                else
                {
                    dao.SaveUser(user);
                    return true;
                }   
            }
            catch (Exception ex)
            {
                throw ex;
            }                
        }

        public MT_NHAN_VIEN getLastUser()
        {   
            MT_NHAN_VIEN LastUser = new MT_NHAN_VIEN();
            try
            {
                LastUser = dao.getLastUser();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return LastUser;
        }

        public bool UpdateUser( MT_NHAN_VIEN user )
        {
            bool isUpdate = false;
            try
            {
               isUpdate = dao.UpdateUser(user);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return isUpdate;
        }

        public bool DelUser( MT_NHAN_VIEN user )
        {
            bool isDeleted = false;
            try
            {
                isDeleted = dao.DeleteUser(user);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return isDeleted;
        }

        public string getGroupUser( string item )
        {
            string groupCode;
            try
            {
                groupCode = dao.getGroupCode(item);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return groupCode;
        }

        public int SaveListUser( List<MT_NHAN_VIEN> listNhanVien )
        {
            try
            {
                return dao.SaveListUser(listNhanVien);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public bool CheckDuplicate( MT_NHAN_VIEN staff )
        {
            try
            {
                if (dao.checkUserDuplicate(staff))
                {
                    return false;
                }
                else
                {      
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
