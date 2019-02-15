using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

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
                MessageBox.Show("Thông số không hợp lệ!");
                return;
            }
            string sqlConnection = "Data Source="+ source + "; Initial Catalog="+database+"; Persist Security Info=True;User ID="+user+"; Password="+pass;
            Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            config.AppSettings.File = "App.config";
            config.AppSettings.Settings["Conn"].Value = sqlConnection;
            config.Save();
            ConfigurationManager.RefreshSection("appSettings");

        }

        private void Config_Load( object sender, EventArgs e )
        {
            Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            config.AppSettings.File = "App.config";

            string connection = config.AppSettings.Settings["Conn"].Value;
            // Data Source=99910DES25\\SQLEXPRESS;Initial Catalog=QL_CONG_TAC_PHI;Persist Security Info=True;User ID=sa;Password=vrb123456"
            if (!string.IsNullOrEmpty(connection))
            {
                //int len = connection.Length;
                //tbSource.Text = connection.Substring(12, connection.IndexOf(@";Initial")-12);
                //tbDataBase.Text = connection.Substring(connection.IndexOf(@"Initial Catalog=")+16, connection.IndexOf(@";Persist") - 50);
                //tbUser.Text = connection.Substring(connection.LastIndexOf(@"User ID="), connection.IndexOf(@";Password"));
                //tbPass.Text = connection.Substring(connection.IndexOf(@"Password="));
                //string abc = "abc";

            }

        }
    }
}
