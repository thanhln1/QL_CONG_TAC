using DAO;
using DTO;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BUS
{
    public class MT_DON_GIA_BUS
    {
        MT_DON_GIA_DAO dao = new MT_DON_GIA_DAO();
        public List<MT_DON_GIA> getDongia(string diachi)
        {
           List<MT_DON_GIA> donGia = new List<MT_DON_GIA>();
            try
            {
                donGia = dao.getDonGia(diachi);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return donGia;
        }

        public List<MT_DON_GIA> getAllDongia()
        {
            List<MT_DON_GIA> donGia = new List<MT_DON_GIA>();
            try
            {
                donGia = dao.getAllDonGia();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return donGia;
        }

        public bool UpdateDonGia(MT_DON_GIA dongia)
        {
            bool isUpdate = false;
            try
            {
                isUpdate = dao.UpdateDonGia(dongia);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return isUpdate;
        }
    }
}
