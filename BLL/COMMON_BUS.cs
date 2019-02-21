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
    }
}
