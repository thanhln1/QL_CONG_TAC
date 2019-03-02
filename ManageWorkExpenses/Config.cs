using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DTO;
using Dapper;

namespace ManageWorkExpenses
{
    public partial class Config : Form
    {
       
        public Config()
        {
            InitializeComponent();
        }

        private void btnSave_Click( object sender, EventArgs e )
        {
            string source = tbSource.Text.Trim();
            string database = tbDataBase.Text.Trim();
            string user = tbUser.Text.Trim();
            string pass = tbPass.Text;

            if (string.IsNullOrEmpty(source) || string.IsNullOrEmpty(user) || string.IsNullOrEmpty(pass))
            {
                MessageBox.Show("Thông số không hợp lệ", "Thông báo !", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string sqlConnection = "Data Source="+ source + "; Initial Catalog="+database+"; Persist Security Info=True;User ID="+user+"; Password="+pass;
            //Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            //config.AppSettings.File = "App.config";
            //config.AppSettings.Settings["Conn"].Value = sqlConnection;
            //config.Save();
            //ConfigurationManager.RefreshSection("appSettings");
            WriteConnectionString(sqlConnection);
            return;

        }

        private void Config_Load( object sender, EventArgs e )
        {
            //Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.PerUserRoamingAndLocal);
            //config.AppSettings.File = "App_manage.config";
            //string connection = config.AppSettings.Settings["conn"].Value;

            string connection = ReadConnectionString();
            txtConnectionString.Text = connection;
            if (string.IsNullOrEmpty(connection))
            {
                txtConnectionString.Text = "Chưa thiết lập kết nối với cơ sở dữ liệu";
                //int len = connection.Length;
                //tbSource.Text = connection.Substring(12, connection.IndexOf(@";Initial")-12);
                //tbDataBase.Text = connection.Substring(connection.IndexOf(@"Initial Catalog=")+16, connection.IndexOf(@";Persist") - 50);
                //tbUser.Text = connection.Substring(connection.LastIndexOf(@"User ID="), connection.IndexOf(@";Password"));
                //tbPass.Text = connection.Substring(connection.IndexOf(@"Password="));
                //string abc = "abc";
            }
        }

        public void WriteConnectionString(string connectionString)
        {
            if (!Directory.Exists("DataSource"))
            {
                Directory.CreateDirectory("DataSource");
            }
            using (StreamWriter sw = new StreamWriter("DataSource/connectionString.txt"))
            {
                sw.WriteLine(connectionString);
                MessageBox.Show("Cài đặt Cở sở dữ liệu thành công - Cần khởi động lại phần mềm", "Thông báo !", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtConnectionString.Text = connectionString;
            }
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
