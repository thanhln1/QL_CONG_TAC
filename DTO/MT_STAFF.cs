using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DTO
{

    public class MT_STAFF
    {
        public string MA_NHAN_VIEN { get; set; }
        public string HO_TEN { get; set; }
        public int SO_NGAY_CONG_TAC { get; set; }
        public List<DateTime> NGAY_CONG_TAC { get; set; }
    }
}
