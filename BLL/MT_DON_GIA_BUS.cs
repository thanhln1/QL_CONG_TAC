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
    }
}
