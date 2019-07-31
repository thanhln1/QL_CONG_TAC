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

        // lấy đơn giá thanh toán công tác phí theo địa điểm
        public double GetDonGia( string diadiem )
        {
            List<MT_DON_GIA> listDonGia = new List<MT_DON_GIA>();
            listDonGia = getDongia(diadiem);
            if (listDonGia == null)
            {
                return 0;
            }            
            return listDonGia[0].DON_GIA;
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
