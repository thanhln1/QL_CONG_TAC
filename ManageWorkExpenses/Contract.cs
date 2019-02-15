using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using DTO;

namespace ManageWorkExpenses
{
    public partial class Contract : Form
    {
        public Contract()
        {
            InitializeComponent();
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            var fileContent = string.Empty;
            var filePath = string.Empty;

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "Excell files (*.xlsx)| Ole Excel File (*.xls)|All files (*.*)|*.*";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //Get the path of specified file
                    filePath = openFileDialog.FileName;

                    //Read the contents of the file into a stream
                    var fileStream = openFileDialog.OpenFile();

                    using (StreamReader reader = new StreamReader(fileStream))
                    {
                        fileContent = reader.ReadToEnd();


                        //Create COM Objects. Create a COM object for everything that is referenced
                        Excel.Application xlApp = new Excel.Application();
                        Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(filePath);
                        Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                        Excel.Range xlRange = xlWorksheet.UsedRange;

                        int rowCount = xlRange.Rows.Count;
                        int colCount = xlRange.Columns.Count +1 ;

                        //iterate over the rows and columns and print to the console as it appears in the file
                        //excel is not zero based!!
                        for (int i = 2; i <= rowCount; i++)
                        {
                            MT_HOP_DONG contract = new MT_HOP_DONG();

                            //write the value to the console 
                            //SO_HOP_DONG
                            xlRange.Cells[i, 1].Value2.ToString();

                            //NGAY_HOP_DONG
                            xlRange.Cells[i, 2].Value2.ToString();

                            //NGAY_THANH_LY
                            xlRange.Cells[i, 3].Value2.ToString();

                            //KHACH_HANG
                            xlRange.Cells[i, 4].Value2.ToString();

                            //MA_KHACH_HANG
                            xlRange.Cells[i, 5].Value2.ToString();

                            //DIA_CHI
                            xlRange.Cells[i, 6].Value2.ToString();

                            //TINH
                            xlRange.Cells[i, 7].Value2.ToString();

                            //GIA_TRI_HOP_DONG
                            xlRange.Cells[i, 8].Value2.ToString();

                            //TONG_CHI_PHI_MUC_TOI_DA
                            xlRange.Cells[i, 9].Value2.ToString();

                            //CHI_PHI_THUC_DA_CHI
                            xlRange.Cells[i, 10].Value2.ToString();

                            //GHI_CHU
                            xlRange.Cells[i, 2].Value2.ToString();

                        }

                        //cleanup
                        GC.Collect();
                        GC.WaitForPendingFinalizers();

                        //  rule of thumb for releasing com objects:
                        //  never use two dots, all COM objects must be referenced and released individually
                        //  ex: [somthing].[something].[something] is bad

                        //release com objects to fully kill excel process from running in the background
                        Marshal.ReleaseComObject(xlRange);
                        Marshal.ReleaseComObject(xlWorksheet);

                        //close and release
                        xlWorkbook.Close();
                        Marshal.ReleaseComObject(xlWorkbook);

                        //quit and release
                        xlApp.Quit();
                        Marshal.ReleaseComObject(xlApp);
                    }
                }
            }

           //  MessageBox.Show(fileContent, "File Content at path: " + filePath, MessageBoxButtons.OK);
        }
    }
}
// gen màu tự đông
//Random random = new Random();
//int randomNumber1 = random.Next(0, 255);
//int randomNumber2 = random.Next(0, 255);
//int randomNumber3 = random.Next(0, 255);
//Color.FromArgb(randomNumber1, randomNumber2, randomNumber3);