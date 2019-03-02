using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

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
            //return ConfigurationManager.ConnectionStrings[name].ConnectionString;
            return ReadConnectionString();
        }

        public string ReadConnectionString()
        {
            string conn = "";

            if (!Directory.Exists("DataSource"))
            {
                return conn;
            }

            string path = Directory.GetCurrentDirectory().ToString() + @"\DataSource\connectionString.txt";
            StreamReader reader = new StreamReader(path, Encoding.UTF8);
            conn = reader.ReadToEnd();
            reader.Close();
            return conn;
        }
    }
}
