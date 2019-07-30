using DAO;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BUS
{
    public class COMMON_BUS
    {
        COMMON daoCommon = new COMMON();

        public static int DAY_OF_WORKING = 3;
        public DateTime ToDateTime(string s, string format = "dd/MM/yyyy", string cultureString = "en-GB")
        {
            try
            {
                var r = DateTime.ParseExact(s, format, CultureInfo.GetCultureInfo(cultureString));
                return r;
            }
            catch (FormatException)
            {
                throw;
            }
            catch (CultureNotFoundException)
            {
                throw; // Given Culture is not supported culture
            }
        }

        public DateTime ToDateTime(string s, string format, CultureInfo culture)
        {
            try
            {
                var r = DateTime.ParseExact(s, format, culture);
                return r;
            }
            catch (FormatException)
            {
                throw;
            }
            catch (CultureNotFoundException)
            {
                throw; // Given Culture is not supported culture
            }

        }

        public bool ResetDB()
        {            
            try
            {
                return daoCommon.ResetDB();
            }
            catch (Exception ex)
            {
                throw ex;
            }
           
        }

        public static DateTime ConverToDateTime( dynamic value )
        {
            DateTime dt = new DateTime();
            // Set value.                               
            if (value != null)
            {
                if (value is double)
                {
                    dt= DateTime.FromOADate((double)value);
                }
                else
                {
                    DateTime.TryParse((string)value, out dt);
                }
            }
           return dt;
        }
    }
}
