using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ManageWorkExpenses
{
    public partial class ProgressForm : Form
    {
        public ProgressForm()
        {
            InitializeComponent();
        }    
        internal void UpdateProgress( int progress )
        {
            if (progressBar.InvokeRequired)
                progressBar.BeginInvoke(new Action(() => progressBar.Value = progress));
            else
                progressBar.Value = progress;  
        }
    }
}
