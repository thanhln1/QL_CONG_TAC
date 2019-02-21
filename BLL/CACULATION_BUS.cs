using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;    

namespace BUS
{
    public class CACULATION_BUS
    {
        // Flag that indcates if a process is running
        private bool isProcessRunning = false;
        public void RunCaculation( int month, int year, string Algorithm )
        {
            // If a process is already running, warn the user and cancel the operation
            if (isProcessRunning)
            {
                MessageBox.Show("Chương trình đang chạy.");
                return;
            }

            // Initialize the dialog that will contain the progress bar
            ProgressForm progressDialog = new ProgressForm();

            // Initialize the thread that will handle the background process
            Thread backgroundThread = new Thread(
                new ThreadStart(() =>
                {
                    // Set the flag that indicates if a process is currently running
                    isProcessRunning = true;

                    // Iterate from 0 - 99
                    // On each iteration, pause the thread for .05 seconds, then update the dialog's progress bar
                    for (int n = 0 ; n < 100 ; n++)
                    {
                        Thread.Sleep(50);
                        progressDialog.UpdateProgress(n);
                    }

                    // Show a dialog box that confirms the process has completed
                    MessageBox.Show("Đã xử lý xong!");

                    // Close the dialog if it hasn't been already
                    if (progressDialog.InvokeRequired)
                        progressDialog.BeginInvoke(new Action(() => progressDialog.Close()));

                    // Reset the flag that indicates if a process is currently running
                    isProcessRunning = false;
                }
            ));

            // Start the background process thread
            backgroundThread.Start();

            // Open the dialog
            progressDialog.ShowDialog();
        }
    }
}

