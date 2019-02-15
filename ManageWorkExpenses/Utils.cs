using System;

namespace ManageWorkExpenses
{
    public static class Utils
    {
        // public static string LogDirectory = System.Configuration.ConfigurationSettings.AppSettings["LOGDIRECTORY"];
        public static string LogDirectory = ".\\";
        public static string LogFilePath = LogDirectory + "Log_" + DateTime.Now.ToString("dd_MM_yyyy__HHmmss") + ".log";  
    }

}