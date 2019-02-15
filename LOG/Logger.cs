using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Configuration;

namespace Log
{
    public class Logger
    {
        private string filePath;
        Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);

        public Logger( string pFilePath )
        {
            filePath = pFilePath;
        }
        public void log( String message )
        {
            string LogMode = config.AppSettings.Settings["DEBUGMODE"].Value;
            if (LogMode.Equals("ON"))
            {
                DateTime datet = DateTime.Now;
                try
                {
                    if (!File.Exists(filePath))
                    {
                        FileStream files = File.Create(filePath);
                        files.Close();
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message.ToString());
                }
                try
                {
                    StreamWriter sw = File.AppendText(filePath);
                    sw.WriteLine(datet.ToString("dd/MM/yyyy hh:mm") + ": " + message);
                    sw.Flush();
                    sw.Close();
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message.ToString());
                }
            }
        }
    }
}
