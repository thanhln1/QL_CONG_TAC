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
        public void setMaxValue(int maxValue)
        {
            if (progressBar.InvokeRequired)
            {
                progressBar.BeginInvoke(new Action(() =>
                {
                    progressBar.Maximum = maxValue;
                }
                ));
            }

        }

        public void SetIndeterminate( bool isIndeterminate )
        {
            if (progressBar.InvokeRequired)
            {
                progressBar.BeginInvoke(new Action(() =>
                {
                    if (isIndeterminate)
                        progressBar.Style = ProgressBarStyle.Marquee;
                    else
                        progressBar.Style = ProgressBarStyle.Blocks;
                }
                ));
            }
            else
            {
                if (isIndeterminate)
                    progressBar.Style = ProgressBarStyle.Marquee;
                else
                    progressBar.Style = ProgressBarStyle.Blocks;
            }               
        }
    }
}
