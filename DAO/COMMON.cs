using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DAO
{
   public class COMMON
    {
        public string LoadConnectionStringOLE()
        {
            string connection = "Data Source=.\\DATABASE.db;Version=3;";
            //string connection = ConfigurationManager.ConnectionStrings["Default"].ConnectionString;
            
            return connection;
        }

        public string LoadConnectionString()
        {  
            string connection = ConfigurationManager.ConnectionStrings["Default"].ConnectionString;

            return connection;
        }

        public string ConnectionString(string name)
        {
            // var connection =   ConfigurationManager.AppSettings["Conn"].ToString();
            return ConfigurationManager.ConnectionStrings[name].ConnectionString;
        }
    }
}
