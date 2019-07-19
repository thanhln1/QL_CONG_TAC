using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DTO
{
    public class OBJ_CALC
    {
        public string MA_NHAN_VIEN { get; set; }
        public List<FROM_TO> LIST_DAY_NOT_WORKING { get; set; }  
        
        public int SO_NGAY_CON_TRONG { get; set; }   

    }

    public class FROM_TO
    {       
        public DateTime FromDate { get; set; }
        public DateTime ToDate { get; set; }
    }
}
