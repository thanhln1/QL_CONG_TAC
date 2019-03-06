using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Data;
using System.Data.SqlClient;
using System.Security.Cryptography;

namespace DAO
{
   
   public class COMMON
    {

        public string ConnectionString( string name )
        {
            // var connection =   ConfigurationManager.AppSettings["Conn"].ToString();
            //return ConfigurationManager.ConnectionStrings[name].ConnectionString;

            // New Connection method
            Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            config.AppSettings.File = "App.config";

            string source = config.AppSettings.Settings["DATASOURCE"].Value;
            string database = config.AppSettings.Settings["DB"].Value;
            string user = config.AppSettings.Settings["USERID"].Value;
            string pass = config.AppSettings.Settings["PASSWORD"].Value;             
            string connection = "Data Source="+source+";Initial Catalog="+database+";Persist Security Info=True;User ID="+user+";Password="+ DecryptString(pass,SECRETKEY);

            // string connection = config.AppSettings.Settings["CONNECTION"].Value;
            return connection;
        }

        public static string SECRETKEY = "NguyenDangThe";

        private static byte[] _salt = Encoding.ASCII.GetBytes("QLCTP_NguyenDangThe");
        public static string EncryptString( string plainText, string sharedSecret )
        {
            if (string.IsNullOrEmpty(plainText))
                throw new ArgumentNullException("plainText");
            if (string.IsNullOrEmpty(sharedSecret))
                throw new ArgumentNullException("sharedSecret");

            string outStr = null;                       // Encrypted string to return
            RijndaelManaged aesAlg = null;              // RijndaelManaged object used to encrypt the data.

            try
            {
                // generate the key from the shared secret and the salt
                Rfc2898DeriveBytes key = new Rfc2898DeriveBytes(sharedSecret, _salt);

                // Create a RijndaelManaged object
                // with the specified key and IV.
                aesAlg = new RijndaelManaged();
                aesAlg.Key = key.GetBytes(aesAlg.KeySize / 8);
                aesAlg.IV = key.GetBytes(aesAlg.BlockSize / 8);

                // Create a decrytor to perform the stream transform.
                ICryptoTransform encryptor = aesAlg.CreateEncryptor(aesAlg.Key, aesAlg.IV);

                // Create the streams used for encryption.
                using (MemoryStream msEncrypt = new MemoryStream())
                {
                    using (CryptoStream csEncrypt = new CryptoStream(msEncrypt, encryptor, CryptoStreamMode.Write))
                    {
                        using (StreamWriter swEncrypt = new StreamWriter(csEncrypt))
                        {

                            //Write all data to the stream.
                            swEncrypt.Write(plainText);
                        }
                    }
                    outStr = Convert.ToBase64String(msEncrypt.ToArray());
                }
            }
            catch
            {
                outStr = plainText;
            }
            finally
            {
                // Clear the RijndaelManaged object.
                if (aesAlg != null)
                    aesAlg.Clear();
            }

            // Return the encrypted bytes from the memory stream.
            return outStr;
        }
        public static string DecryptString( string cipherText, string sharedSecret )
        {
            if (string.IsNullOrEmpty(cipherText))
                throw new ArgumentNullException("cipherText");
            if (string.IsNullOrEmpty(sharedSecret))
                throw new ArgumentNullException("sharedSecret");

            // Declare the RijndaelManaged object
            // used to decrypt the data.
            RijndaelManaged aesAlg = null;

            // Declare the string used to hold
            // the decrypted text.
            string plaintext = null;

            try
            {
                // generate the key from the shared secret and the salt
                Rfc2898DeriveBytes key = new Rfc2898DeriveBytes(sharedSecret, _salt);

                // Create a RijndaelManaged object
                // with the specified key and IV.
                aesAlg = new RijndaelManaged();
                aesAlg.Key = key.GetBytes(aesAlg.KeySize / 8);
                aesAlg.IV = key.GetBytes(aesAlg.BlockSize / 8);

                // Create a decrytor to perform the stream transform.
                ICryptoTransform decryptor = aesAlg.CreateDecryptor(aesAlg.Key, aesAlg.IV);
                // Create the streams used for decryption.                
                byte[] bytes = Convert.FromBase64String(cipherText);
                using (MemoryStream msDecrypt = new MemoryStream(bytes))
                {
                    using (CryptoStream csDecrypt = new CryptoStream(msDecrypt, decryptor, CryptoStreamMode.Read))
                    {
                        using (StreamReader srDecrypt = new StreamReader(csDecrypt))

                            // Read the decrypted bytes from the decrypting stream
                            // and place them in a string.
                            plaintext = srDecrypt.ReadToEnd();
                    }
                }
            }
            catch
            {
                plaintext = "";
            }
            finally
            {
                // Clear the RijndaelManaged object.
                if (aesAlg != null)
                    aesAlg.Clear();
            }

            return plaintext;
        }

        #region hàm cũ không sử dụng
        public string LoadConnectionStringOLE()
        {
            string connection = "Data Source=.\\DATABASE.db;Version=3;";
            //string connection = ConfigurationManager.ConnectionStrings["Default"].ConnectionString;
            
            return connection;
        }

        public string LoadConnectionString()
        {  
            // Old Connection
            // string connection = ConfigurationManager.ConnectionStrings["Default"].ConnectionString;

            // New Connection
            Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);                   
            config.AppSettings.File = "App.config";
            string connection = config.AppSettings.Settings["CONNECTION"].Value;              

            return connection;
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
        #endregion
    }
}
