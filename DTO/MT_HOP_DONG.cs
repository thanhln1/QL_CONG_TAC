using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DTO
{
   public class MT_HOP_DONG
    {
        public int ID { get; set; }
        public string SO_HOP_DONG { get; set; }
        public DateTime NGAY_HOP_DONG { get; set; }
        public DateTime NGAY_THANH_LY { get; set; }
        public string KHACH_HANG { get; set; }
        public string MA_KHACH_HANG { get; set; }  
        public string NHOM_KHACH_HANG { get; set; }
        public string DIA_CHI { get; set; }
        public string TINH { get; set; }
        public int GIA_TRI_HOP_DONG { get; set; }
        public int TONG_CHI_PHI_MUC_TOI_DA { get; set; }
        public int CHI_PHI_THUC_DA_CHI { get; set; }
        public string GHI_CHU { get; set; }

    }
}
