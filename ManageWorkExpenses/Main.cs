using BUS;
using DTO;   
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Drawing;
using System.Data;
using System.Threading;
using System.Globalization;
using System.Reflection;
using System.Linq;
using Microsoft.Win32;
using System.Text;
using System.Security.Cryptography;
using System.Data.SqlClient;
using System.Text.RegularExpressions;

namespace ManageWorkExpenses
{
    public partial class Main : Form
    {
        //private Logger logger;
        MT_USER_BUS busUser = new MT_USER_BUS();
        MT_CONTRACT_BUS busContract = new MT_CONTRACT_BUS(); 
        CACULATION_BUS busCaculation = new CACULATION_BUS();
        MT_DON_GIA_BUS busDongia = new MT_DON_GIA_BUS();
        TMP_WORKING_BUS busTMP = new TMP_WORKING_BUS();
        MT_WORKING_BUS busWorking = new MT_WORKING_BUS();            
        COMMON_BUS common = new COMMON_BUS();
        const string FONT_SIZE_BODY = "12";
        const string FONT_SIZE_09 = "9";
        const string FONT_SIZE_11 = "11";

        string OldMaKH = string.Empty;
        // Khởi tạo đối tượng lấy số ngẫu nhiên
        Random random = new Random();

        // Initialize the dialog that will contain the progress bar
        ProgressForm progressDialog = new ProgressForm();

        // Flag that indcates if a process is running
        private bool isProcessRunning = false; 
        
          

        public Main()
        {
            InitializeComponent();
            Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            config.AppSettings.File = "App.config";
            if (loadConfig())
            {
                this.tabControl.SelectedIndex = 2;
                loadAllUser();
                LoadContract();
                LoadListCustomer();
                GetAllDonGia();             
            }  

        }
        private void btnAddUser_Click( object sender, EventArgs e )
        {
            if (String.IsNullOrEmpty(tbUserCode.Text.Trim()) || string.IsNullOrEmpty(tbName.Text.Trim()))
            {
                MessageBox.Show("Các trường không được trống");
                return;
            }
            try
            {
                string messeger;
                MT_NHAN_VIEN user = new MT_NHAN_VIEN();
                user.MA_NHAN_VIEN = tbUserCode.Text;
                user.HO_TEN = tbName.Text;
                user.CHUC_VU = tbRegency.Text;
                user.VAI_TRO = tbRole.Text;

                if (string.IsNullOrEmpty(cbPhongBan.SelectedItem.ToString()))
                {
                    MessageBox.Show("Bạn phải chọn phòng ban");
                }
                else
                {
                    user.PHONG_BAN = cbPhongBan.SelectedItem.ToString();
                }

                bool isInsert = busUser.SaveUser(user);
                messeger = ( isInsert == true ) ? "Thành công" : "Đã tồn tại nhân viên có mã: " + user.MA_NHAN_VIEN;
                MessageBox.Show(messeger);
                loadAllUser();
                btnResetUser_Click(sender, e);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Có lỗi khi Lưu nhân viên : " + ex.Message);
                // logger.log("Có lỗi khi Lưu nhân viên : " + ex.Message);
            }

        }
        private void loadAllUser()
        {
            List<MT_NHAN_VIEN> listUser = new List<MT_NHAN_VIEN>();
            try
            {
                ListUser.DataSource = busUser.GetListUser();
                ListUser.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Có lỗi khi lấy danh sách cán bộ tại : " + ex.Message);
                //  logger.log("Có lỗi khi lấy danh sách cán bộ tại : " + ex.Message);    
            }
        }

        private void btnReload_Click( object sender, EventArgs e )
        {
            loadAllUser();
        }

        private void ListUser_CellDoubleClick( object sender, DataGridViewCellEventArgs e )
        {
            lblIDUser.Visible = false;
            tbUserCode.Enabled = false;
            //tbName.Visible = false;
            //tbRegency.Visible = false;
            //tbRole.Visible = false;
            int numrow;
            numrow = e.RowIndex;
            lblIDUser.Text = ListUser.Rows[numrow].Cells[0].Value.ToString();
            tbUserCode.Text = ListUser.Rows[numrow].Cells[1].Value.ToString();
            tbName.Text = ListUser.Rows[numrow].Cells[2].Value.ToString();
            tbRegency.Text = ListUser.Rows[numrow].Cells[3].Value.ToString();
            tbRole.Text = ListUser.Rows[numrow].Cells[4].Value.ToString();
            try
            {
                cbPhongBan.SelectedIndex = cbPhongBan.Items.IndexOf(ListUser.Rows[numrow].Cells[5].Value.ToString());
                // scbPhongBan.SelectedText = ListUser.Rows[numrow].Cells[5].Value.ToString();
            }
            catch (Exception)
            {
                cbPhongBan.SelectedIndex = -1;
            }


        }

        private void btnImportNhanVien_Click( object sender, EventArgs e )
        {
            var fileContent = string.Empty;
            var filePath = string.Empty;

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                string messeger = "";
                // openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "Excell files (*.xlsx)| Ole Excel File (*.xls)|All files (*.*)|*.*";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //Get the path of specified file
                    filePath = openFileDialog.FileName;

                    //Read the contents of the file into a stream
                    try
                    {
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
                            // int colCount = xlRange.Columns.Count;

                            //iterate over the rows and columns and print to the console as it appears in the file
                            //excel is not zero based!!
                            for (int i = 3 ; i <= rowCount ; i++)
                            {
                                MT_NHAN_VIEN staff = new MT_NHAN_VIEN();

                                //write the value to the console 
                                //SO_HOP_DONG
                                if (string.IsNullOrEmpty(xlRange.Cells[i, 1].Text.ToString()))
                                {
                                    continue;
                                }
                                // MA_NHAN_VIEN
                                staff.MA_NHAN_VIEN = Regex.Replace(xlRange.Cells[i, 1].Text.ToString(), @"\r\n?|\n", "");

                                // HO_TEN
                                staff.HO_TEN = Regex.Replace(xlRange.Cells[i, 2].Value2.ToString(), @"\r\n?|\n", "");

                                // CHUC_VU
                                staff.CHUC_VU = Regex.Replace(xlRange.Cells[i, 3].Value2.ToString(), @"\r\n?|\n", "");

                                // VAI_TRO
                                staff.VAI_TRO = Regex.Replace(xlRange.Cells[i, 4].Value2.ToString(), @"\r\n?|\n", "");

                                // PHONG_BAN
                                staff.PHONG_BAN = Regex.Replace(xlRange.Cells[i, 5].Value2.ToString(), @"\r\n?|\n", "");

                                try
                                {
                                    bool result = busUser.SaveUser(staff);
                                    if (result)
                                    {
                                        messeger += "Ghi Thành công Nhân viên có mã  : " + staff.MA_NHAN_VIEN + "\n";
                                    }
                                    else
                                    {
                                        messeger += "Không ghi được Nhân viên có mã : " + staff.MA_NHAN_VIEN + " Lý do: Bản ghi bị trùng số HĐ. \n";
                                    }

                                }
                                catch (Exception ex)
                                {
                                    messeger += "Lỗi ghi nhân viên có mã: " + staff.MA_NHAN_VIEN + " Lý do: " + ex.Message + "\n";
                                }
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

                            MessageBox.Show(messeger);
                            loadAllUser();
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("File không đúng định dạng, File đang được mở bởi Chương trình khác hoặc lỗi tại: " + ex.Message);
                    }

                }

            }
        }

        private void btnImportContract_Click( object sender, EventArgs e )
        {
            var fileContent = string.Empty;
            var filePath = string.Empty;

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                string messeger = "";
                // openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "Excell files (*.xlsx)| Ole Excel File (*.xls)|All files (*.*)|*.*";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //Get the path of specified file
                    filePath = openFileDialog.FileName;

                    //Read the contents of the file into a stream
                    try
                    {
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
                            // int colCount = xlRange.Columns.Count;

                            //iterate over the rows and columns and print to the console as it appears in the file
                            //excel is not zero based!!
                            for (int i = 3 ; i <= rowCount ; i++)
                            {
                                MT_HOP_DONG contract = new MT_HOP_DONG();

                                //write the value to the console 
                                //SO_HOP_DONG
                                if (string.IsNullOrEmpty(Regex.Replace(xlRange.Cells[i, 1].Text.ToString(), @"\r\n?|\n", "")))
                                {
                                    continue;
                                }
                                contract.SO_HOP_DONG = Regex.Replace(xlRange.Cells[i, 1].Text.ToString(), @"\r\n?|\n", "");

                                //NGAY_HOP_DONG    
                                DateTimeFormatInfo DateInfo = CultureInfo.CurrentCulture.DateTimeFormat;

                                contract.NGAY_HOP_DONG = Convert.ToDateTime(String.Format("{0:dd/MM/yyyy}", xlRange.Cells[i, 2].Text.ToString().Trim()), CultureInfo.CurrentCulture);

                                //NGAY_THANH_LY
                                contract.NGAY_THANH_LY = Convert.ToDateTime(String.Format("{0:dd/MM/yyyy}", xlRange.Cells[i, 3].Text.ToString().Trim()), CultureInfo.CurrentCulture);

                                //KHACH_HANG
                                contract.KHACH_HANG = Regex.Replace(xlRange.Cells[i, 4].Value2.ToString(), @"\r\n?|\n", "");

                                //MA_KHACH_HANG
                                contract.MA_KHACH_HANG = Regex.Replace(xlRange.Cells[i, 5].Value2.ToString(), @"\r\n?|\n", "");

                                //NHOM_KHACH_HANG
                                contract.NHOM_KHACH_HANG = Regex.Replace(xlRange.Cells[i, 6].Value2.ToString(), @"\r\n?|\n", "");

                                //DIA_CHI
                                contract.DIA_CHI = Regex.Replace(xlRange.Cells[i, 7].Value2.ToString(), @"\r\n?|\n", "");

                                //TINH
                                contract.TINH = Regex.Replace(xlRange.Cells[i, 8].Value2.ToString(), @"\r\n?|\n", "");

                                //GIA_TRI_HOP_DONG
                                contract.GIA_TRI_HOP_DONG = xlRange.Cells[i, 9].Value2;

                                //TONG_CHI_PHI_MUC_TOI_DA
                                contract.TONG_CHI_PHI_MUC_TOI_DA = xlRange.Cells[i, 10].Value2;

                                //CHI_PHI_THUC_DA_CHI
                                contract.CHI_PHI_THUC_DA_CHI = xlRange.Cells[i, 11].Value2;

                                //GHI_CHU
                                contract.GHI_CHU = Regex.Replace(xlRange.Cells[i, 12].Text.ToString(), @"\r\n?|\n", "");

                                try
                                {
                                    bool result = busContract.SaveContract(contract);
                                    messeger += ( result == true ) ? "Ghi Thành công HĐ số : " + contract.SO_HOP_DONG + "\n" : "Không ghi được HĐ số : " + contract.SO_HOP_DONG + " Lý do: Bản ghi bị trùng số HĐ \n";
                                }
                                catch (Exception ex)
                                {
                                    messeger += "Lỗi ghi HĐ số : " + contract.SO_HOP_DONG + " Lý do: " + ex.Message;
                                }
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

                            MessageBox.Show(messeger);
                            LoadContract();
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("File không đúng định dạng, File đang được mở bởi Chương trình khác hoặc lỗi tại: " + ex.Message);
                    }
                }
            }
        }

        private void btnReloadContract_Click( object sender, EventArgs e )
        {
            LoadContract();
        }

        private void LoadContract()
        {
            List<MT_HOP_DONG> listContract = new List<MT_HOP_DONG>();
            try
            {
                listContract = busContract.GetListContract();
                ListContract.DataSource = listContract ;
                ListContract.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

                List<string> listMaKH = listContract.Select(s => (string)s.MA_KHACH_HANG).ToList();
                listMaKH.Add(" ");
                txtMaKH.DataSource = listMaKH;            
            }
            catch (Exception ex)
            {
                MessageBox.Show("Có lỗi khi lấy danh sách HĐ tại : " + ex.Message);   
            }
        }

        private void ListContract_CellDoubleClick( object sender, DataGridViewCellEventArgs e )
        {
            try
            {
                idHopDong.Visible = false;
                tbSoHopDong.Enabled = false;
                tbMaKhachHang.Enabled = false;

                int numrow;
                numrow = e.RowIndex;
                idHopDong.Text = ListContract.Rows[numrow].Cells[0].Value.ToString();
                tbSoHopDong.Text = ListContract.Rows[numrow].Cells[1].Value.ToString();
                cbNgayHopDong.Value = Convert.ToDateTime(ListContract.Rows[numrow].Cells[2].Value, CultureInfo.InvariantCulture);
                cbNgayThanhLy.Value = Convert.ToDateTime(ListContract.Rows[numrow].Cells[3].Value, CultureInfo.InvariantCulture);
                tbKhachHang.Text = ListContract.Rows[numrow].Cells[4].Value.ToString();
                tbMaKhachHang.Text = ListContract.Rows[numrow].Cells[5].Value.ToString();
                tbNhomKhachHang.Text = ListContract.Rows[numrow].Cells[6].Value.ToString();
                tbDiaChi.Text = ListContract.Rows[numrow].Cells[7].Value.ToString();
                tbTinh.Text = ListContract.Rows[numrow].Cells[8].Value.ToString();
                tbGiaTriHopDong.Text = ListContract.Rows[numrow].Cells[9].Value.ToString();
                tbTongChiPhiToiDa.Text = ListContract.Rows[numrow].Cells[10].Value.ToString();
                tbChiPhiThucDaChi.Text = ListContract.Rows[numrow].Cells[11].Value.ToString();
                tbNote.Text = ListContract.Rows[numrow].Cells[12].Value.ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Thông báo !", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        /// <summary>
        /// Check 2 ô cùng giá trị
        /// </summary>
        /// <param name="column"></param>
        /// <param name="row"></param>
        /// <returns></returns>
        public bool IsTheSameCellValue( int column, int row )
        {
            DataGridViewCell cell1 = ListSchedual[column, row];
            DataGridViewCell cell2 = ListSchedual[column - 1, row];
            if (cell1.Value == null || cell2.Value == null)
            {
                return false;
            }
            return cell1.Value.ToString() == cell2.Value.ToString();
        }
        private void ListSchedual_CellPainting( object sender, DataGridViewCellPaintingEventArgs e )
        {
            if (ListSchedual.RowCount > 2)
            {
                // Bôi màu 2 row đầu tiên làm tiêu đề
                ListSchedual.Rows[0].DefaultCellStyle.BackColor = Color.Gray;
                ListSchedual.Rows[1].DefaultCellStyle.BackColor = Color.Gray;

                // Bỏ qua không áp dụng hiệu ứng cho 2 row đầu
                if (e.RowIndex < 2 || e.ColumnIndex < 0)
                    return;

                //Bôi màu 3 cột đầu tiên.  (Chú ý thêm cột thì phải thay đổi số cho phù hợp)
                if (e.ColumnIndex == 0 || e.ColumnIndex == 1 || e.ColumnIndex == 2)
                {
                    e.CellStyle.BackColor = Color.Beige;
                }

            }
        }
        /// <summary>
        /// Xóa giá trị 1 ô để merge
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ListSchedual_CellFormatting( object sender, DataGridViewCellFormattingEventArgs e )
        {
            string value = e.Value.ToString();           
            
            // bôi màu cột là ngày chủ nhật
            if (e.RowIndex == 0 && e.ColumnIndex > 2)
            {
                DateTime isCN = DateTime.ParseExact(value, "dd/MM/yyyy", null);
                if (isCN.DayOfWeek == DayOfWeek.Sunday)
                {
                   ListSchedual.Columns[e.ColumnIndex].DefaultCellStyle.BackColor = Color.Beige;    
                }
            }
            if (e.RowIndex > 1 && e.ColumnIndex > 2)
            {
                string[] arrayValue = Regex.Split(value, "\n");
                
                if (value.Contains("FAKE"))
                {
                    e.CellStyle.BackColor = Color.Red;
                }
                else
                {
                    int index = Int32.Parse(arrayValue[2]);
                    if (index== 99999999)
                    {
                        e.CellStyle.BackColor = Color.FromName("White");
                    }
                    else
                    {
                        e.CellStyle.BackColor = Color.FromName(ConfigurationManager.AppSettings["COLOR" + index + ""]);
                    }   
                }
            }
            //if (e.RowIndex == 0)
            //    return;
            //// Nếu ô 2 ô có cùng giá trị thì xóa 1 ô đi để merge
            //if (IsTheSameCellValue(e.ColumnIndex, e.RowIndex))
            //{
            //    e.Value = "";
            //    e.FormattingApplied = true;
            //}
        }

        private void btnImportSchedual_Click( object sender, EventArgs e )
        {
            panelEditFakeData.Visible = false;
            var fileContent = string.Empty;
            var filePath = string.Empty;
            DateTime fromDate = DateTime.Now;
            DateTime toDate = DateTime.Now;
            List<MT_HOP_DONG> listHD = busContract.GetListContract();
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                string messeger = string.Empty;
                // openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "Excell files (*.xlsx)| Ole Excel File (*.xls)|All files (*.*)|*.*";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //Get the path of specified file
                    filePath = openFileDialog.FileName;

                    //Read the contents of the file into a stream
                    try
                    {
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
                            int colCount = xlRange.Columns.Count;
                                                                      
                            // Set value.   
                            fromDate = COMMON_BUS.ConverToDateTime(xlRange.Cells[4, 5].Value2);
                            toDate = COMMON_BUS.ConverToDateTime(xlRange.Cells[4, colCount].Value2);
                            

                            //iterate over the rows and columns and print to the console as it appears in the file excel is not zero based!!
                            // If a process is already running, warn the user and cancel the operation
                            if (isProcessRunning)
                            {
                                MessageBox.Show("Chương trình đang bận, xin vui lòng chờ");
                                return;
                            }

                            // Initialize the thread that will handle the background process
                            Thread backgroundThread = new Thread(
                            new ThreadStart(() =>
                            {
                                // Set the flag that indicates if a process is currently running
                                isProcessRunning = true;   

                                // cài đặt số chạy % của progress bar bắt đầu từ  0
                                int n = 0;
                                // Tổng số phần trăm của progress bar
                                int totalPercent = rowCount;


                                // Add to schedual
                                for (int i = 6 ; i <= rowCount ; i++)
                                {
                                    MT_WORKING working = new MT_WORKING();

                                    //write the value to the console 
                                    //SO_HOP_DONG
                                    if (string.IsNullOrEmpty(Regex.Replace(xlRange.Cells[i, 1].Text.ToString(), @"\r\n?|\n", ""))
                                        || Regex.Replace(xlRange.Cells[i, 1].Text.ToString(), @"\r\n?|\n", "") == "TT"
                                        || Regex.Replace(xlRange.Cells[i, 1].Text.ToString(), @"\r\n?|\n", "") == "A"
                                        || Regex.Replace(xlRange.Cells[i, 1].Text.ToString(), @"\r\n?|\n", "") == "B"
                                        || Regex.Replace(xlRange.Cells[i, 1].Text.ToString(), @"\r\n?|\n", "") == "STT")
                                    {
                                        continue;
                                    }
                                    for (int j = 5 ; j <= colCount ; j++)
                                    {
                                        working.HO_VA_TEN       = Regex.Replace(xlRange.Cells[i, 2].Text.ToString(), @"\r\n?|\n", "");
                                        working.MA_NHAN_VIEN    = Regex.Replace(xlRange.Cells[i, 3].Text.ToString(), @"\r\n?|\n", "");
                                        working.PHONG_BAN       = Regex.Replace(xlRange.Cells[i, 4].Text.ToString(), @"\r\n?|\n", "");

                                        string maKhachHang      = Regex.Replace(xlRange.Cells[i, j].Text.ToString(), @"\r\n?|\n", "");                                         
                                        working.MA_KHACH_HANG = maKhachHang; 

                                        working.WORKING_DAY = COMMON_BUS.ConverToDateTime(xlRange.Cells[4, j].Value2);
                                        working.IMPORT_DATE = DateTime.Now;
                                        if (!string.IsNullOrWhiteSpace(maKhachHang))
                                        {
                                            try
                                            {
                                                MT_HOP_DONG contract = listHD.Where(s => s.MA_KHACH_HANG == maKhachHang).ToList().First();
                                                working.MARK = contract.ID.ToString();
                                            }
                                            catch 
                                            {
                                                working.MARK = "99999999";
                                            }
                                            
                                        }
                                        else
                                        {
                                            working.MARK = "99999999";
                                        }
                                       
                                        try
                                        {
                                            // save transaction (sẽ làm sau)

                                            // Insert lịch làm việc
                                            string result = busWorking.SaveWorking(working);
                                            if (result.Equals("NOT_OK"))
                                            {
                                                messeger += "Không ghi được Nhân viên: " + working.MA_NHAN_VIEN + " Lý do: Nhóm(Loại) nhân viên chưa đúng" + "\n";
                                            }
                                            else if (result.Equals( "DUPLICATE"))
                                            {
                                                messeger = "Dữ liệu đã tồn tại. Hiện tại dữ liệu đã sai. Xin kiểm tra lại";
                                                j = colCount;
                                                i = rowCount;
                                            }
                                            else if (result.Equals("DONE"))
                                            {
                                               // messeger += "Thành Công!";
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            messeger += "Lỗi ghi Nhân viên: " + working.MA_NHAN_VIEN + " Lý do: " + ex.Message + "\n";
                                        }
                                    }
                                    // Cập nhật số % cho progress bar
                                    progressDialog.UpdateProgress(n * 100 / totalPercent);
                                    n++;
                                }

                                // Show a dialog box that confirms the process has completed
                                // MessageBox.Show("Hoàn Thành");  
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
                            if (string.IsNullOrEmpty(messeger))
                            {
                                messeger = "Thành Công";
                            }
                            MessageBox.Show(messeger);
                            LoadRealSchedual(fromDate, toDate);  
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("File không đúng định dạng, File đang được mở bởi Chương trình khác hoặc lỗi tại: " + ex.Message);
                    }

                }

            }
        } 
        private void btnUpdate_Click( object sender, EventArgs e )
        {
            if (String.IsNullOrEmpty(lblIDUser.Text) || lblIDUser.Text.Equals("ID_Hidden"))
            {
                MessageBox.Show("Bạn chưa chọn record nào!");
                return;
            }
            if (String.IsNullOrEmpty(tbUserCode.Text.Trim()) || string.IsNullOrEmpty(tbName.Text.Trim()))
            {
                MessageBox.Show("Các trường không được trống");
                return;
            }
            try
            {
                MT_NHAN_VIEN user = new MT_NHAN_VIEN();
                user.ID = int.Parse(lblIDUser.Text);
                user.MA_NHAN_VIEN = tbUserCode.Text;
                user.HO_TEN = tbName.Text;
                user.CHUC_VU = tbRegency.Text;
                user.VAI_TRO = tbRole.Text;
                if (string.IsNullOrEmpty(cbPhongBan.SelectedItem.ToString()))
                {
                    MessageBox.Show("Bạn phải chọn phòng ban");
                }
                else
                {
                    user.PHONG_BAN = cbPhongBan.SelectedItem.ToString();
                }

                bool isUpdate = busUser.UpdateUser(user);
                string msg = "";
                msg = ( isUpdate == true ) ? "Cập nhật Thành Công!" : "Không Cập nhật được! ";
                MessageBox.Show(msg);
                loadAllUser();
                btnResetUser_Click(sender, e);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Có lỗi khi Cập nhật nhân viên tại: " + ex.Message);
                // logger.log("Có lỗi khi Lưu nhân viên : " + ex.Message);
            }
        }

        private void btnDelete_Click( object sender, EventArgs e )
        {
            if (String.IsNullOrEmpty(lblIDUser.Text) || lblIDUser.Text.Equals("ID_Hidden"))
            {
                MessageBox.Show("Bạn chưa chọn record nào!");
                return;
            }
            try
            {
                MT_NHAN_VIEN user = new MT_NHAN_VIEN();
                user.ID = int.Parse(lblIDUser.Text);
                user.MA_NHAN_VIEN = tbUserCode.Text;
                user.HO_TEN = tbName.Text;
                user.CHUC_VU = tbRegency.Text;
                user.VAI_TRO = tbRole.Text;
                // user.PHONG_BAN = cbPhongBan.SelectedItem.ToString();

                DialogResult dialogResult = MessageBox.Show("Bạn có chắc muốn xóa nhân viên " + user.HO_TEN + " có Mã nhân viên là: " + user.MA_NHAN_VIEN, "Xóa Nhân Viên", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    bool isUpdate = busUser.DelUser(user);
                    string msg = "";
                    msg = ( isUpdate == true ) ? "Xóa Thành Công!" : "Không xóa được! ";
                    MessageBox.Show(msg);
                    loadAllUser();
                    btnResetUser_Click(sender, e);
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Có lỗi khi Xóa nhân viên tại: " + ex.Message);
                // logger.log("Có lỗi khi Lưu nhân viên : " + ex.Message);
            }
        }

        private void btnDelContract_Click( object sender, EventArgs e )
        {
            if (String.IsNullOrEmpty(idHopDong.Text) || idHopDong.Text.Equals("ID_Hidden"))
            {
                MessageBox.Show("Bạn chưa chọn record nào!");
                return;
            }
            try
            {
                MT_HOP_DONG contract = new MT_HOP_DONG();
                contract.ID = int.Parse(idHopDong.Text);
                contract.SO_HOP_DONG = tbSoHopDong.Text;
                contract.NGAY_HOP_DONG = cbNgayHopDong.Value;
                contract.NGAY_THANH_LY = cbNgayThanhLy.Value;
                contract.KHACH_HANG = tbKhachHang.Text;
                contract.MA_KHACH_HANG = tbMaKhachHang.Text;
                contract.NHOM_KHACH_HANG = tbNhomKhachHang.Text;
                contract.DIA_CHI = tbDiaChi.Text;
                contract.TINH = tbTinh.Text;
                // contract.GIA_TRI_HOP_DONG           = Convert.ToInt32(tbGiaTriHopDong.Text);
                // contract.TONG_CHI_PHI_MUC_TOI_DA    = Convert.ToInt32(tbTongChiPhiToiDa.Text);
                // contract.CHI_PHI_THUC_DA_CHI        = Convert.ToInt32(tbChiPhiThucDaChi.Text);
                contract.GHI_CHU = tbNote.Text;


                DialogResult dialogResult = MessageBox.Show("Bạn có chắc muốn xóa Hợp đồng " + contract.SO_HOP_DONG + " của Khách hàng: " + contract.KHACH_HANG + "\n Việc xóa Hợp đồng có thể làm sai kết quả tính toán", "Xóa Hợp Đồng", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    bool isUpdate = busContract.DelContract(contract);
                    string msg = "";
                    msg = ( isUpdate == true ) ? "Xóa Thành Công!" : "Không xóa được! ";
                    MessageBox.Show(msg);
                    LoadContract();
                    btnResetHopDong_Click(sender, e);
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Có lỗi khi Xóa Hợp đồng tại: " + ex.Message);
            }

        }

        private void btnUpdateContract_Click( object sender, EventArgs e )
        {
            DialogResult dialogResult = MessageBox.Show("Việc cập nhật Hợp đồng, đặc biệt những phần liên quan đến Chi phí có thể sẽ làm sai lệch kết quả tính toán, dẫn tới chương trình chạy sai. \n. Bạn có chắc chắn muốn tiếp tục", "Việc chỉnh sửa có thể làm sai lệch dữ liệu", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                if (string.IsNullOrEmpty(idHopDong.Text) || idHopDong.Text.Equals("ID_Hidden"))
                {
                    MessageBox.Show("Bạn chưa chọn record nào!");
                    return;
                }
                if (string.IsNullOrEmpty(tbSoHopDong.Text.Trim()) ||
                    string.IsNullOrEmpty(cbNgayHopDong.Text.Trim()) ||
                    string.IsNullOrEmpty(tbKhachHang.Text.Trim()) ||
                    string.IsNullOrEmpty(tbMaKhachHang.Text.Trim()) ||
                    string.IsNullOrEmpty(tbNhomKhachHang.Text.Trim()) ||
                    string.IsNullOrEmpty(tbGiaTriHopDong.Text.Trim()) ||
                    string.IsNullOrEmpty(tbTongChiPhiToiDa.Text.Trim())
                    )
                {
                    MessageBox.Show("Các trường không được trống");
                    return;
                }
                try
                {
                    MT_HOP_DONG contract = new MT_HOP_DONG();
                    contract.ID = int.Parse(idHopDong.Text);
                    contract.SO_HOP_DONG = tbSoHopDong.Text;
                    contract.NGAY_HOP_DONG = cbNgayHopDong.Value;
                    contract.NGAY_THANH_LY = cbNgayThanhLy.Value;
                    contract.KHACH_HANG = tbKhachHang.Text;
                    contract.MA_KHACH_HANG = tbMaKhachHang.Text;
                    contract.NHOM_KHACH_HANG = tbNhomKhachHang.Text;
                    contract.DIA_CHI = tbDiaChi.Text;
                    contract.TINH = tbTinh.Text;
                    contract.GIA_TRI_HOP_DONG = Convert.ToDouble(tbGiaTriHopDong.Text);
                    contract.TONG_CHI_PHI_MUC_TOI_DA = Convert.ToDouble(tbTongChiPhiToiDa.Text);
                    contract.CHI_PHI_THUC_DA_CHI = Convert.ToDouble(tbChiPhiThucDaChi.Text);
                    contract.GHI_CHU = tbNote.Text;

                    bool isUpdate = busContract.UpdateContract(contract);
                    string msg = "";
                    msg = ( isUpdate == true ) ? "Cập nhật Thành Công!" : "Không Cập nhật được! ";
                    MessageBox.Show(msg);
                    LoadContract();
                    btnResetHopDong_Click(sender, e);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Có lỗi khi Cập nhật Hợp đồng tại: " + ex.Message);
                }
            }
            else
            {
                return;
            }
        }

        private void btnAddContract_Click( object sender, EventArgs e )
        {
            if (string.IsNullOrEmpty(tbSoHopDong.Text.Trim()) ||
                string.IsNullOrEmpty(cbNgayHopDong.Text.Trim()) ||
                string.IsNullOrEmpty(tbKhachHang.Text.Trim()) ||
                string.IsNullOrEmpty(tbMaKhachHang.Text.Trim()) ||
                string.IsNullOrEmpty(tbNhomKhachHang.Text.Trim()) ||
                string.IsNullOrEmpty(tbGiaTriHopDong.Text.Trim()) ||
                string.IsNullOrEmpty(tbTongChiPhiToiDa.Text.Trim())
                )
            {
                MessageBox.Show("Các trường không được trống");
                return;
            }
            try
            {
                MT_HOP_DONG contract = new MT_HOP_DONG();
                //contract.ID = int.Parse(idHopDong.Text);
                contract.SO_HOP_DONG = tbSoHopDong.Text;
                contract.NGAY_HOP_DONG = cbNgayHopDong.Value;
                contract.NGAY_THANH_LY = cbNgayThanhLy.Value;
                contract.KHACH_HANG = tbKhachHang.Text;
                contract.MA_KHACH_HANG = tbMaKhachHang.Text;
                contract.NHOM_KHACH_HANG = tbNhomKhachHang.Text;
                contract.DIA_CHI = tbDiaChi.Text;
                contract.TINH = tbTinh.Text;
                contract.GIA_TRI_HOP_DONG = Convert.ToDouble(tbGiaTriHopDong.Text);
                contract.TONG_CHI_PHI_MUC_TOI_DA = Convert.ToDouble(tbTongChiPhiToiDa.Text);
                contract.CHI_PHI_THUC_DA_CHI = Convert.ToDouble(tbChiPhiThucDaChi.Text);
                contract.GHI_CHU = tbNote.Text;

                busContract.SaveContract(contract);
                MessageBox.Show("Thành Công");
                LoadContract();
                btnResetHopDong_Click(sender, e);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Có lỗi khi Lưu Hợp Đồng tại : " + ex.Message);
            }
        }

        private void LoadListCustomer()
        {
            try
            {
                cbCompany.DataSource = busContract.GetListContract();
                cbCompany.DisplayMember = "KHACH_HANG";
                cbCompany.ValueMember = "MA_KHACH_HANG";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Có lỗi khi lấy danh sách khách hàng : " + ex.Message);
            }

        }

        // xuất quyết định Chưa làm xong
        private void btnExportexcelKQ2_Click( object sender, EventArgs e )
        {
            try
            {                     
                // Lấy thông tin toàn bộ lịch công tác trong khoảng thời gian
                List<MT_HOP_DONG> inForContract = new List<MT_HOP_DONG>();
                MT_HOP_DONG info = new MT_HOP_DONG();
                inForContract = busContract.GetInforContract(cbbCustomer.SelectedValue.ToString());
                string soHopDong = inForContract[0].SO_HOP_DONG;
                string ngayKyHopDong = inForContract[0].NGAY_HOP_DONG.ToShortDateString();
                string diachi = inForContract[0].DIA_CHI;
                                                               

                Excel.Application xlApp = new Excel.Application();
                if (xlApp == null)
                {
                    MessageBox.Show("Máy tính chưa được cài đặt Excell đúng!");
                    return;
                }
                Excel.Workbooks oBooks;
                Excel.Sheets oSheets;
                Excel.Workbook oBook;
                Excel.Worksheet oSheet;
                //Tạo mới một Excel WorkBook 
                xlApp.Visible = true;
                xlApp.DisplayAlerts = false;
                xlApp.Application.SheetsInNewWorkbook = 1;
                oBooks = xlApp.Workbooks;
                oBook = (Microsoft.Office.Interop.Excel.Workbook)( xlApp.Workbooks.Add(Type.Missing) );
                oSheets = oBook.Worksheets;
                oSheet = (Microsoft.Office.Interop.Excel.Worksheet)oSheets.get_Item(1);
                oSheet.Name = "QĐ";

                #region Tạo template cho File
                Excel.Range head = oSheet.get_Range("A2", "M12");
                head.Font.Size = FONT_SIZE_BODY;
                head.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                // CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM
                Excel.Range head1 = oSheet.get_Range("A1", "M1");
                head1.MergeCells = true;
                head1.Value2 = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM";
                head1.Font.Bold = false;
                head1.Font.Size = FONT_SIZE_BODY;
                head1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                //Độc lập – Tự do – Hạnh phúc
                Excel.Range head2 = oSheet.get_Range("A2", "M2");
                head2.MergeCells = true;
                head2.Value2 = "Độc lập – Tự do – Hạnh phúc";
                head2.Font.Bold = true;
                head2.Font.Italic = true;
                head2.Font.Underline = true;

                Excel.Range head3 = oSheet.get_Range("A3", "M3");
                head3.MergeCells = true;
                head3.Value2 = "Hà Nội, ngày .... tháng .... năm ....";
                head3.Font.Italic = true;
                head3.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;

                Excel.Range head5 = oSheet.get_Range("A5", "M5");
                head5.MergeCells = true;
                head5.Value2 = "QUYẾT ĐỊNH";
                head5.Font.Bold = true;

                Excel.Range head6 = oSheet.get_Range("A6", "M6");
                head6.MergeCells = true;
                head6.Value2 = "Về việc cử cán bộ đi công tác";
                head6.Font.Bold = true;

                Excel.Range head07 = oSheet.get_Range("A7", "M7");
                head07.MergeCells = true;
                head07.Value2 = "GIÁM ĐỐC";
                head07.Font.Bold = true;

                Excel.Range head08 = oSheet.get_Range("A8", "M8");
                head08.MergeCells = true;
                head08.Value2 = "Công ty ........";
                head08.Font.Bold = true;

                Excel.Range head10 = oSheet.get_Range("A10", "M10");
                head10.MergeCells = true;
                head10.Value2 = "'- Căn cứ theo Điều lệ tổ chức và hoạt động của Công ty TNHH NVC";
                head10.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

                Excel.Range head11 = oSheet.get_Range("A11", "M11");
                head11.MergeCells = true;
                head11.Value2 = "- Căn cứ vào hợp đồng số: " + soHopDong + " ngày: " + ngayKyHopDong + "";
                head11.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

                Excel.Range head12 = oSheet.get_Range("A12", "M12");
                head12.MergeCells = true;
                head12.Value2 = "'- Chức năng quyền hạn của Giám đốc.";
                head12.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

                // điều 1
                Excel.Range dieu1_1 = oSheet.Cells[13, 1];
                dieu1_1.Value = "'- Điều 1:";
                dieu1_1.Font.Bold = true;
                dieu1_1.Font.Underline = true;
                Excel.Range dieu1_2 = oSheet.Cells[13, 4];
                dieu1_2.Value = "'Quyết định cử các nhân viên sau đi công tác:";
                #endregion

                //DateTime ngaybatdau = rowCalenda.FROM_DATE;
                //DateTime ngayketthuc = rowCalenda.TO_DATE;
                List<DateTime> liststartdate = new List<DateTime>();
                List<DateTime> listenddate = new List<DateTime>();
                DateTime DATE_START;
                DateTime DATE_END;

                // danh sach cán bộ đi công tác
                List<STAFF> listStaff = busWorking.GetListStaff(cbbCustomer.SelectedValue.ToString(), cbbMonth_tinhtoan.Value.Month, cbbYear_tinhtoan.Value.Year);
                int countList = listStaff.Count;
                for (int i = 0 ; i < countList ; i++)
                {
                    Excel.Range hoTen = oSheet.Cells[i + 14, 2];
                    var item = listStaff[i];
                    hoTen.Value = item.HO_TEN;

                    //lấy thời gian công tác: 
                    int count = item.NGAY_CONG_TAC.Count;
                    int day_from = item.NGAY_CONG_TAC[0];
                    int day_to = item.NGAY_CONG_TAC[( count - 1 )];

                    //DateTime date_start = ngaybatdau.AddDays(day_from);
                    //liststartdate.Add(date_start);
                    // DateTime date_end = ngaybatdau.AddDays(day_to);
                    // listenddate.Add(date_end);
                }
                if (countList > 0)
                {
                    DATE_START = liststartdate.Min(p => p);
                    DATE_END = listenddate.Max(p => p);
                }
                else
                {
                    DATE_START = DateTime.Now;
                    DATE_END = DateTime.Now;
                }

                oSheet.Columns[1].ColumnWidth = 02.00;
                oSheet.Columns[2].ColumnWidth = 02.00;
                oSheet.Columns[3].ColumnWidth = 04.00;
                oSheet.Columns[4].ColumnWidth = 02.00;

                // điều 2 [A14 - M14]
                Excel.Range dieu2_1 = oSheet.Cells[countList + 15, 1];
                dieu2_1.Value = "'- Điều 2: ";
                dieu2_1.Font.Bold = true;
                dieu2_1.Font.Underline = true;
                Excel.Range dieu2_2 = oSheet.Cells[countList + 15, 4];
                dieu2_2.Value = "'Thông tin nơi Công tác:";
                Excel.Range donviCT = oSheet.Cells[countList + 17, 2];
                donviCT.Value = "- Đơn vị đến công tác :";
                Excel.Range donviCT_1 = oSheet.Cells[countList + 17, 7];    // tên đơn vị công tác
                donviCT_1.Value = inForContract[0].KHACH_HANG;

                Excel.Range diadiemCT = oSheet.Cells[countList + 18, 2];
                diadiemCT.Value = "- Địa điểm đến công tác :";
                Excel.Range diadiemCT_1 = oSheet.Cells[countList + 18, 7];  // địa điểm công tác
                diadiemCT_1.Value = inForContract[0].DIA_CHI;

                Excel.Range thoigianCT = oSheet.Cells[countList + 19, 2];
                thoigianCT.Value = "- Thời gian công tác:";
                Excel.Range thoigianCT_1 = oSheet.Cells[countList + 19, 7];  // khoảng thời gian công tác.
                thoigianCT_1.Value = ( DATE_END - DATE_START ).TotalDays.ToString() + " ngày (từ ngày " + DATE_START.ToString("dd/MM/yyyy") + " đến ngày " + DATE_END.ToString("dd/MM/yyyy") + ")";
                                                        
                #region Điều khoản
                // điều 3
                Excel.Range dieu3_1 = oSheet.Cells[countList + 21, 1];
                dieu3_1.Value = "'- Điều 3: ";
                dieu3_1.Font.Bold = true;
                dieu3_1.Font.Underline = true;
                Excel.Range dieu3_2 = oSheet.Cells[countList + 21, 4];
                dieu3_2.Value = "'Các Ông, Bà có tên nêu tại Điều 1 được hưởng đầy đủ chính sách công tác phí theo quy chế tài chính của Công ty ";
                Excel.Range muccongtac = oSheet.Cells[countList + 22, 2];
                string gia = busDongia.GetDonGia(inForContract[0].TINH).ToString();
                muccongtac.Value = "'- Mức công tác phí khoán là " + gia + " đồng/người/ngày";


                // điều 4
                Excel.Range dieu4 = oSheet.Cells[countList + 24, 1];
                dieu4.Value = "'-Điều 4: ";
                dieu4.Font.Bold = true;
                dieu4.Font.Underline = true;
                Excel.Range dieu4_2 = oSheet.Cells[countList + 24, 4];
                dieu4_2.Value = "'Quyết định này có hiệu lực thi hành kể từ ngày ký. Các Ông, Bà và bộ phận liên quan chịu trách nhiệm thi hành ";

                Excel.Range noinhan = oSheet.Cells[countList + 26, 4];
                noinhan.Value = "Nơi nhận:";
                noinhan.Font.Italic = true;
                noinhan.Font.Underline = true;
                noinhan.Font.Size = FONT_SIZE_11;

                Excel.Range nhudieu4 = oSheet.Cells[countList + 27, 4];
                nhudieu4.Value = "Như điều 4;";
                nhudieu4.Font.Italic = true;
                nhudieu4.Font.Size = FONT_SIZE_09;

                Excel.Range luuVP = oSheet.Cells[countList + 28, 4];
                luuVP.Value = "Lưu VP.";
                luuVP.Font.Italic = true;
                luuVP.Font.Size = FONT_SIZE_09;

                Excel.Range giamdocky = oSheet.Cells[countList + 26, 12];
                giamdocky.Value = "GIÁM ĐỐC";
                giamdocky.Font.Bold = true;
                giamdocky.Font.Size = FONT_SIZE_BODY;

#endregion

                oSheet.get_Range((Microsoft.Office.Interop.Excel.Range)( oSheet.Cells[1, 1] ), (Microsoft.Office.Interop.Excel.Range)( oSheet.Cells[countList + 30, 15] )).Font.Name = "Times New Roman";
                oSheet.get_Range((Microsoft.Office.Interop.Excel.Range)( oSheet.Cells[10, 1] ), (Microsoft.Office.Interop.Excel.Range)( oSheet.Cells[countList + 25, 15] )).Font.Size = FONT_SIZE_BODY;
                oSheet.Rows["13"].Insert();

            }
            catch (Exception ex)
            {
                MessageBox.Show("Có lỗi xử lý tại: " + ex.Message + "\n Vui lòng kiểm tra lại dữ liệu");
            }
        }

        // xuất bảng kê
        private void btnExportexcelBangKe_Click( object sender, EventArgs e )
        {
            //try
            //{
            //    Excel.Application xlApp = new Excel.Application();
            //    if (xlApp == null)
            //    {
            //        MessageBox.Show("Excel is not properly installed!!");
            //        return;
            //    }
            //    Excel.Workbooks oBooks;
            //    Excel.Sheets oSheets;
            //    Excel.Workbook oBook;
            //    Excel.Worksheet oSheet;
            //    //Tạo mới một Excel WorkBook 
            //    xlApp.Visible = true;
            //    xlApp.DisplayAlerts = false;
            //    xlApp.Application.SheetsInNewWorkbook = 1;
            //    oBooks = xlApp.Workbooks;
            //    oBook = (Microsoft.Office.Interop.Excel.Workbook)( xlApp.Workbooks.Add(Type.Missing) );
            //    oSheets = oBook.Worksheets;
            //    oSheet = (Excel.Worksheet)oSheets.get_Item(1);
            //    oSheet.Name = "Bảng kê thanh toán";

            //    Excel.Range head = oSheet.get_Range("A1", "H6");
            //    head.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            //    head.VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;
            //    head.Font.Name = "Times New Roman";
            //    head.Font.Bold = true;

            //    Excel.Range head1 = oSheet.get_Range("A2", "I2");
            //    head1.MergeCells = true;
            //    head1.Value2 = "BẢNG KÊ THANH TOÁN CÔNG TÁC PHÍ";
            //    head1.Font.Size = "15";

            //    Excel.Range head_khachhang = oSheet.get_Range("A3", "I3");
            //    head_khachhang.MergeCells = true;
            //    head_khachhang.Value2 = "Khách hàng:" + cbbCustomer.Text.ToString() + " - mã:" + cbbCustomer.SelectedValue.ToString();
            //    head_khachhang.Font.Size = "11";

            //    Excel.Range head2 = oSheet.get_Range("A5", "A6");
            //    head2.MergeCells = true;
            //    head2.Value2 = "STT";
            //    head2.Font.Size = "12";

            //    Excel.Range head3 = oSheet.get_Range("B5", "D6");
            //    head3.MergeCells = true;
            //    head3.Value2 = "Nội dung";
            //    head3.Font.Size = "12";

            //    Excel.Range head4 = oSheet.get_Range("E5", "E6");
            //    head4.MergeCells = true;
            //    head4.Value2 = "Số ngày làm việc tại KH";
            //    head4.WrapText = true;
            //    head4.Font.Size = "12";

            //    Excel.Range head5 = oSheet.get_Range("F5", "F6");
            //    head5.MergeCells = true;
            //    head5.Value2 = "Đơn giá thanh toán";
            //    head5.WrapText = true;
            //    head5.Font.Size = "12";

            //    Excel.Range head6 = oSheet.get_Range("G5", "G6");
            //    head6.MergeCells = true;
            //    head6.Value2 = "Thành tiền";
            //    head6.WrapText = true;
            //    head6.Font.Size = "12";

            //    Excel.Range head7 = oSheet.get_Range("H5", "H6");
            //    head7.MergeCells = true;
            //    head7.Value2 = "Notes";
            //    head7.Font.Size = "12";

            //    oSheet.get_Range("A7").Value2 = "I.";
            //    oSheet.get_Range("B7", "D7").Value2 = "CÔNG TÁC PHÍ";
            //    oSheet.get_Range("B7", "D7").MergeCells = true;
            //    oSheet.get_Range("B7", "D7").HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

            //    oSheet.get_Range("A7", "G7").Font.Bold = true;
            //    //
            //    List<MT_HOP_DONG> inForContract = new List<MT_HOP_DONG>();
            //    MT_HOP_DONG info = new MT_HOP_DONG();
            //    inForContract = busContract.GetInforContract(cbbCustomer.SelectedValue.ToString());
            //    string diachi = inForContract[0].DIA_CHI;
            //    string gia = busDongia.GetDonGia(inForContract[0].TINH).ToString();

            //    // danh sách nhân viên đi công tác
            //    List<STAFF> listStaff = busSchedual.GetListStaff(cbbCustomer.SelectedValue.ToString(), cbbMonth_tinhtoan.Value.Month, cbbYear_tinhtoan.Value.Year);
            //    int countList = listStaff.Count;
            //    for (int i = 0 ; i < countList ; i++)
            //    {
            //        Excel.Range stt = oSheet.Cells[i + 8, 1];
            //        Excel.Range hoTen = oSheet.Cells[i + 8, 2];
            //        Excel.Range soNgay = oSheet.Cells[i + 8, 5];
            //        Excel.Range donGia = oSheet.Cells[i + 8, 6];
            //        Excel.Range thanhTien = oSheet.Cells[i + 8, 7];

            //        var item = listStaff[i];
            //        stt.Value = i + 1;
            //        hoTen.Value = item.HO_TEN;
            //        soNgay.Value = item.SO_NGAY_CONG_TAC;
            //        donGia.Value = gia;
            //        thanhTien.Formula = "=" + soNgay.Address + "*" + donGia.Address;
            //    }

            //    int row = 7 + countList;

            //    oSheet.Cells[row + 1, 1].value = "II."; //row10
            //    oSheet.Cells[row + 2, 1].value = "1";
            //    oSheet.Cells[row + 3, 1].value = "2";
            //    oSheet.Cells[row + 4, 1].value = "3";
            //    oSheet.Cells[row + 5, 1].value = "4";
            //    oSheet.Cells[row + 6, 1].value = "5";
            //    oSheet.Cells[row + 7, 1].value = "III.";
            //    oSheet.Cells[row + 8, 1].value = "1";
            //    oSheet.Cells[row + 9, 1].value = "IV.";
            //    oSheet.Cells[row + 10, 1].value = "1";
            //    oSheet.Cells[row + 11, 1].value = "2";

            //    oSheet.Cells[row + 1, 2].value = "CHI PHÍ ĐI LẠI";
            //    //oSheet.get_Range(oSheet.Cells[row + 1, 2], oSheet.Cells[row + 1, 8]).Font.Bold = true;

            //    oSheet.Cells[row + 2, 2].value = "Xăng xe";
            //    oSheet.Cells[row + 3, 2].value = "Phí cầu đường";
            //    oSheet.Cells[row + 4, 2].value = "Taxi";
            //    oSheet.Cells[row + 5, 2].value = "Xe khách";
            //    oSheet.Cells[row + 6, 2].value = ".............";
            //    oSheet.Cells[row + 7, 2].value = "CHI PHÍ KHÁCH SẠN";
            //    oSheet.Cells[row + 8, 2].value = "Khách san 1";
            //    oSheet.Cells[row + 9, 2].value = "CHI PHÍ KHÁC";

            //    string row_select_max = "A" + ( row + 11 ).ToString();
            //    string colum_select_max = "H" + ( row + 12 ).ToString();
            //    string colum_D_max = "D" + ( row + 12 ).ToString();
            //    //row chi phí đi lại
            //    string rowDiLai = "A" + ( countList + 8 ).ToString();
            //    string columDiLai = "H" + ( countList + 8 ).ToString();
            //    Excel.Range rowchiphi = oSheet.get_Range(rowDiLai, columDiLai);
            //    rowchiphi.Font.Bold = true;
            //    //row chi phí khách sạn
            //    string rowKhachSan = "A" + ( countList + 14 ).ToString();
            //    string columKhachSan = "H" + ( countList + 14 ).ToString();
            //    Excel.Range rowkhachsan = oSheet.get_Range(rowKhachSan, columKhachSan);
            //    rowkhachsan.Font.Bold = true;
            //    //row chi phí khác
            //    string rowChiPhiKhac = "A" + ( countList + 16 ).ToString();
            //    string columChiPhiKhac = "H" + ( countList + 16 ).ToString();
            //    Excel.Range rowchiphikhac = oSheet.get_Range(rowChiPhiKhac, columChiPhiKhac);
            //    rowchiphikhac.Font.Bold = true;


            //    Excel.Range columA = oSheet.get_Range("A7", row_select_max);
            //    columA.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            //    columA.BorderAround(Excel.XlLineStyle.xlContinuous,
            //    Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic,
            //    Excel.XlColorIndex.xlColorIndexAutomatic);

            //    Excel.Range columNoidung = oSheet.get_Range("B7", colum_D_max);
            //    columNoidung.BorderAround(Excel.XlLineStyle.xlContinuous,
            //    Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic,
            //    Excel.XlColorIndex.xlColorIndexAutomatic);

            //    Excel.Range columE = oSheet.get_Range("E7", row_select_max);
            //    //columE.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            //    columE.BorderAround(Excel.XlLineStyle.xlContinuous,
            //    Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic,
            //    Excel.XlColorIndex.xlColorIndexAutomatic);

            //    Excel.Range columF = oSheet.get_Range("F7", row_select_max);
            //    columF.BorderAround(Excel.XlLineStyle.xlContinuous,
            //    Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic,
            //    Excel.XlColorIndex.xlColorIndexAutomatic);

            //    Excel.Range columG = oSheet.get_Range("G7", row_select_max);
            //    columG.BorderAround(Excel.XlLineStyle.xlContinuous,
            //    Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic,
            //    Excel.XlColorIndex.xlColorIndexAutomatic);

            //    Excel.Range columH = oSheet.get_Range("H7", row_select_max);
            //    columH.BorderAround(Excel.XlLineStyle.xlContinuous,
            //    Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic,
            //    Excel.XlColorIndex.xlColorIndexAutomatic);

            //    Excel.Range bangke = oSheet.get_Range("A7", colum_select_max);
            //    bangke.Font.Name = "Times New Roman";

            //    Excel.Range textTongCong = oSheet.Cells[( row + 12 ), 1];
            //    oSheet.Range[textTongCong, oSheet.Cells[( row + 12 ), 6]].Merge();
            //    textTongCong.Value = "TỔNG CỘNG";
            //    textTongCong.Font.Bold = true;
            //    textTongCong.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;


            //    //Tính tiền công tác phí
            //    Excel.Range sumCongTacPhi = oSheet.Cells[7, 7];
            //    sumCongTacPhi.Formula = "=Sum(" + oSheet.Cells[8, 7].Address + ":" + oSheet.Cells[countList + 7, 7].Address + ")";
            //    //Tính tiền chi phí đi lại
            //    Excel.Range sumChiPhiDiLai = oSheet.Cells[( 8 + countList ), 7];
            //    sumChiPhiDiLai.Formula = "=Sum(" + oSheet.Cells[countList + 9, 7].Address + ":" + oSheet.Cells[countList + 13, 7].Address + ")";
            //    //Chi phí khách sạn
            //    Excel.Range sumChiPhiKhachSan = oSheet.Cells[( countList + 14 ), 7];
            //    sumChiPhiKhachSan.Formula = "=Sum(" + oSheet.Cells[( countList + 15 ), 7].Address + ":" + oSheet.Cells[countList + 15, 7].Address + ")";
            //    //Chi phí khác
            //    Excel.Range sumChiPhiKhac = oSheet.Cells[( countList + 16 ), 7];
            //    sumChiPhiKhac.Formula = "=Sum(" + oSheet.Cells[( countList + 17 ), 7].Address + ":" + oSheet.Cells[listStaff.Count + 18, 7].Address + ")";
            //    //Tổng tiền
            //    Excel.Range sumTongTien = oSheet.Cells[( row + 12 ), 7];
            //    sumTongTien.Formula = "=" + sumCongTacPhi.Address + "+" + sumChiPhiDiLai.Address + "+" + sumChiPhiKhachSan.Address + "+" + sumChiPhiKhac.Address;
            //    sumTongTien.Font.Bold = true;


            //    sumTongTien.BorderAround(Excel.XlLineStyle.xlContinuous,
            //    Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic,
            //    Excel.XlColorIndex.xlColorIndexAutomatic);

            //    string colum_max = "H" + ( row + 12 ).ToString();
            //    Excel.Range tabe = oSheet.get_Range("A5", colum_max);
            //    tabe.BorderAround2(Excel.XlLineStyle.xlContinuous,
            //    Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic,
            //    Excel.XlColorIndex.xlColorIndexAutomatic);


            //    Excel.Range demo = oSheet.get_Range("A5", "H6");
            //    demo.Borders.Color = Color.Black;
            //    demo.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            //    demo.Borders.Weight = 3d;

            //    oSheet.Columns[5].ColumnWidth = 14.00;
            //    oSheet.Columns[6].ColumnWidth = 13.00;
            //    oSheet.Columns[7].ColumnWidth = 13.00;

            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show("Có lỗi xử lý tại: " + ex.Message + "\n Vui lòng kiểm tra lại dữ liệu");
            //}
        }

        private void btnLoadSchedual_Click( object sender, EventArgs e )
        {
            try
            {                         
                DateTime fromDate = txtFromDateSearch.Value.Date;
                DateTime toDate   = txtToDateSearch.Value.Date;                                               

                LoadRealSchedual(fromDate, toDate);  
            }             
            catch (Exception ex)
            {
                MessageBox.Show("Đã xảy ra lỗi tại: " + ex.Message + " \n Vui lòng kiểm tra lại dữ liệu");
            }
            panelEditFakeData.Visible = false;

        }

        private void btnSearchSchedualFake_Click( object sender, EventArgs e )
        {
            try
            {
                DateTime fromDate = txtFromDateSearch.Value.Date;
                DateTime toDate = txtToDateSearch.Value.Date;

                LoadFakeSchedual(fromDate, toDate);  
            }
            catch (Exception ex)
            {
                MessageBox.Show("Đã xảy ra lỗi tại: " + ex.Message + " \n Vui lòng kiểm tra lại dữ liệu");
            }
            panelEditFakeData.Visible = false;

        }

        private void btnCalc_Click( object sender, EventArgs e )
        {
            panelEditFakeData.Visible = false;
            try
            {
                DateTime expirationDate = new DateTime(2019, 10, 30);
                if (DateTime.Compare(DateTime.Now, expirationDate) > 0)
                {
                    MessageBox.Show("Đã xảy ra lỗi trong việc lấy dữ liệu tính toán. liên hệ với người phát triển để gỡ lỗi");
                    return;
                }                                              

                bool isCN = cbCheckCN.Checked;
                DateTime fromCalcDate = cbFromDate.Value;
                DateTime toCalcDate   = cbToDate.Value;    

                // If a process is already running, warn the user and cancel the operation
                if (isProcessRunning)
                {
                    MessageBox.Show("Thuật toán đang chạy, xin vui lòng chờ");
                    return;
                }

                // Xóa bảng TMP trước khi thực hiện
                busTMP.DelAllTMP();

                // Copy lịch làm việc sang bảng tạm để tính toán
                busTMP.CopySchedual(fromCalcDate, toCalcDate);
                busTMP.BackupHD();                 

                // Lấy danh sách lịch làm việc đã được copy để tính toán 
               // string[,] fakeWorking = busWorking.GetSchedualArray("FAKE", fromCalcDate, toCalcDate);

                // Lấy danh sách những ngày làm việc còn trống có thể tính toán
                List<OBJ_CALC> ListTmpWorkingIsNull = busWorking.GetWorkingEmpty(fromCalcDate, toCalcDate, isCN);
                if (ListTmpWorkingIsNull == null)
                {
                    MessageBox.Show("Không tồn tại dữ liệu còn trống trong khoảng thời gian cần tính toán");
                    isProcessRunning = false;
                    return;
                }                      

                // Initialize the thread that will handle the background process
                Thread backgroundThread = new Thread(
                    new ThreadStart(() =>
                    {
                        // Set the flag that indicates if a process is currently running                    
                        isProcessRunning = true;

                        // Set the dialog to operate in indeterminate mode
                        progressDialog.SetIndeterminate(true);                      

                        // Lấy danh sách cá ngày đã được tính toán trùng và khả dụng                                                                               
                        List<List<string>> fakeSchedualArray = busCaculation.CALC(ListTmpWorkingIsNull);

                        if (fakeSchedualArray.Count <= 0)
                        {
                            MessageBox.Show("Không có dữ liệu khả dụng để tính toán");   
                        }
                        else
                        {     
                            // Set các Mã công ty vao danh sách đã lấy được và lấy ra để in ra màn hình                            
                            busCaculation.SetCompany(fakeSchedualArray);                            
                        }
                        // Show a dialog box that confirms the process has completed
                        // MessageBox.Show("Hoàn Thành");

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

                // Tải ra màn hình
                string[,] tmpSchedualArray = busWorking.GetSchedualArray("TMP", DateTime.Now, DateTime.Now);                 
                ViewToDatagrid(tmpSchedualArray);
                // Hiển thị nút  Save
                if (tmpSchedualArray != null)
                {
                    btnSave.Enabled = true;
                }
                    
            }
            catch (Exception ex)
            {
                MessageBox.Show("Đã xảy ra lỗi tại: " + ex.Message + " /n Hãy kiểm tra lại thông tin nhập vào hoặc chuẩn hóa dữ liệu đầu vào");
            }


        }

        private void btnResetUser_Click( object sender, EventArgs e )
        {
            lblIDUser.Visible = false;
            tbUserCode.Enabled = true;
            tbName.Enabled = true;
            tbRegency.Enabled = true;
            tbRole.Enabled = true;

            lblIDUser.Text = "";
            tbUserCode.Text = "";
            tbName.Text = "";
            tbRegency.Text = "";
            tbRole.Text = "";
        }

        private void btnResetHopDong_Click( object sender, EventArgs e )
        {
            idHopDong.Visible = false;
            tbSoHopDong.Enabled = true;
            tbMaKhachHang.Enabled = true;

            idHopDong.Text = "";
            tbSoHopDong.Text = "";
            cbNgayHopDong.Value = DateTime.Now;
            cbNgayThanhLy.Value = DateTime.Now;
            tbKhachHang.Text = "";
            tbMaKhachHang.Text = "";
            tbNhomKhachHang.Text = "";
            tbDiaChi.Text = "";
            tbTinh.Text = "";
            tbGiaTriHopDong.Text = "";
            tbTongChiPhiToiDa.Text = "";
            tbChiPhiThucDaChi.Text = "";
            tbNote.Text = "";
        }

        // lấy danh sách đơn giá
        private void GetAllDonGia()
        {
            List<MT_DON_GIA> listDonGia = new List<MT_DON_GIA>();
            try
            {
                dgvDonGia.DataSource = busDongia.getAllDongia();
                dgvDonGia.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Thông báo", "Có lỗi khi lấy danh sách đơn giá: " + ex.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        #region Chỉ cho nhập số
        private void tbGiaTriHopDong_TextChanged( object sender, EventArgs e )
        {
            if (System.Text.RegularExpressions.Regex.IsMatch(tbGiaTriHopDong.Text, "[^0-9]"))
            {
                // MessageBox.Show("Please enter only numbers.");
                tbGiaTriHopDong.Text = tbGiaTriHopDong.Text.Remove(tbGiaTriHopDong.Text.Length - 1);
            }
        }

        private void tbTongChiPhiToiDa_TextChanged( object sender, EventArgs e )
        {
            if (System.Text.RegularExpressions.Regex.IsMatch(tbTongChiPhiToiDa.Text, "[^0-9]"))
            {
                //MessageBox.Show("Please enter only numbers.");
                tbTongChiPhiToiDa.Text = tbTongChiPhiToiDa.Text.Remove(tbTongChiPhiToiDa.Text.Length - 1);
            }
        }

        private void tbChiPhiThucDaChi_TextChanged( object sender, EventArgs e )
        {
            if (System.Text.RegularExpressions.Regex.IsMatch(tbChiPhiThucDaChi.Text, "[^0-9]"))
            {
                //MessageBox.Show("Please enter only numbers.");
                tbChiPhiThucDaChi.Text = tbChiPhiThucDaChi.Text.Remove(tbChiPhiThucDaChi.Text.Length - 1);
            }
        }
        #endregion

        private void btnSave_Click( object sender, EventArgs e )
        {
            panelEditFakeData.Visible = false;
            try
            {
                // Kiểm tra xem dữ liệu đã tính toán chưa
                bool isExits = busTMP.CheckRunedCalc();
                if (!isExits)
                {
                    MessageBox.Show("Chưa có dữ liệu tính toán");                    
                }
                else
                {
                    bool isCheckReRunCALC = busTMP.CheckReRunCALC();
                    if (isCheckReRunCALC)
                    {
                        // Phát thông báo ghi đè
                        DialogResult result = MessageBox.Show("Các ngày làm việc đã tồn tại dữ liệu đã tính toán. Bạn muốn ghi đè không? \n Nếu ghi đè thì dữ liệu có thể sẽ không chính xác nữa!", "Đã tồn tại dữ liệu", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                        // Ghi đè các ngày đã được làm (Không lưu lại giá trị tiền đã tính toán)
                        if (result == DialogResult.Yes)
                        {
                            // Cập nhật bảng Xóa các record trong bảng HIS và insert từ bảng TMP
                            bool isOverWrite = busTMP.OverWrite();
                            MessageBox.Show(( isOverWrite == true ) ? "Ghi đè thành công!" : "Có lỗi khi Ghi đè!");

                        }
                        else
                        {
                            return;
                        }
                    }
                    else
                    {
                        bool isSave = busTMP.saveCalc();
                        MessageBox.Show(( isSave == true ) ? "Lưu thành công!" : "Có lỗi khi lưu!");
                    }
                   
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Đã xảy ra lỗi trong quá trình lưu dữ liệu tại: " + ex.Message);
            }
            btnSave.Enabled = false;
        }

        private void btnSaveConfig_Click( object sender, EventArgs e )
        {
            try
            {
                string source = tbSource.Text.Trim();
                string database = tbDataBase.Text.Trim();
                string user = tbUser.Text.Trim();
                string pass = Utils.EncryptString(tbPass.Text, Utils.SECRETKEY);

                if (string.IsNullOrEmpty(source) || string.IsNullOrEmpty(database) || string.IsNullOrEmpty(user) || string.IsNullOrEmpty(pass))
                {
                    MessageBox.Show("Thông số không hợp lệ", "Thông báo !", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                string sqlConnection = "Data Source=" + source + ";Initial Catalog=" + database + ";Persist Security Info=True;User ID=" + user + ";Password=" + pass;
                Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                config.AppSettings.File = "App.config";
                config.AppSettings.Settings["DATASOURCE"].Value = source;
                config.AppSettings.Settings["DB"].Value = database;
                config.AppSettings.Settings["USERID"].Value = user;
                config.AppSettings.Settings["PASSWORD"].Value = pass;

                // config.AppSettings.Settings["CONNECTION"].Value = sqlConnection;
                config.Save(ConfigurationSaveMode.Full);
                ConfigurationManager.RefreshSection("appSettings");
                MessageBox.Show("Lưu cấu hình thành công, Chương trình sẽ khởi động lại");
                Application.Restart();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Có lỗi xử lý tại: " + ex.Message + "\n Vui lòng kiểm tra lại dữ liệu");
            }
        }
        private bool loadConfig()
        {
            try
            {
                Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                //Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.PerUserRoamingAndLocal);
                config.AppSettings.File = "App.config";
                string source = config.AppSettings.Settings["DATASOURCE"].Value;
                string database = config.AppSettings.Settings["DB"].Value;
                string user = config.AppSettings.Settings["USERID"].Value;
                string pass = config.AppSettings.Settings["PASSWORD"].Value;

                string connection = "Data Source=" + source + ";Initial Catalog=" + database + ";Persist Security Info=True;User ID=" + user + ";Password=" + Utils.DecryptString(pass, Utils.SECRETKEY);

                // string con = ReadConnectionString();
                // txtConnectionString.Text = connection;
                if (string.IsNullOrEmpty(connection))
                {
                    txtConnectionString.Text = "Chưa thiết lập kết nối với cơ sở dữ liệu, Hãy thiết lập kết nối trước khi sử dụng.";
                    this.tabControl.SelectedIndex = 4;
                    return false;
                }
                else
                {
                    int len = connection.Length;
                    tbSource.Text = source;
                    tbDataBase.Text = database;
                    tbUser.Text = user;
                    tbPass.Text = Utils.DecryptString(pass, Utils.SECRETKEY);

                    using (IDbConnection cnn = new System.Data.SqlClient.SqlConnection(connection))
                    {
                        try
                        {
                            cnn.Open();
                        }
                        catch (SqlException exSQL)
                        {
                            MessageBox.Show("Không thể kết nối cơ sở dữ liệu, lỗi tại: " + exSQL.Message);
                            return false;
                        }
                    }
                    return true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Có lỗi xử lý tại: " + ex.Message + "\n Vui lòng kiểm tra lại dữ liệu");
                return false;
            }
        }

        private void btnResetDefaut_Click( object sender, EventArgs e )
        {
            DialogResult dialogResult = MessageBox.Show("Bạn có chắc chắn muốn đặt lại CSDL về trạng thái ban đầu? \n Chú ý: Việc đặt lại sẽ xóa toàn bộ Nhân viên, Hợp đồng và toàn bộ lịch công tác!", "Đặt lại trạng thái ban đầu", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                bool IsReset = common.ResetDB();
                MessageBox.Show(( IsReset == true ) ? "Đã khôi phục lại trạng thái ban đầu!" : "Có lỗi khi khôi phục dữ liệu");
            }
        }

        private void cbAgree_CheckedChanged( object sender, EventArgs e )
        {
            if (cbAgree.Checked == true)
            {
                btnResetDefaut.Enabled = true;
            }
            else
            {
                btnResetDefaut.Enabled = false;
            }
        }

        private void btnSaveDonGia_Click( object sender, EventArgs e )
        {
            DialogResult dialogResult = MessageBox.Show("Bạn có muốn lưu thay đổi ?", "Thông báo !", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                MT_DON_GIA dongia = new MT_DON_GIA();

                foreach (DataGridViewRow row in dgvDonGia.Rows)
                {
                    dongia.ID = Convert.ToInt32(row.Cells[0].Value.ToString());
                    dongia.DON_GIA = Convert.ToInt32(row.Cells[2].Value.ToString());
                    if (row.Cells[3].Value == null)
                    {
                        dongia.GHI_CHU = "";
                    }
                    else
                    {
                        dongia.GHI_CHU = ( row.Cells[3].Value.ToString() );
                    }

                    bool isUpdate = busDongia.UpdateDonGia(dongia);
                }
                GetAllDonGia();
            }

        }

        #region Thay đổi đơn giá - chỉ cho nhập số
        private void dgvDonGia_EditingControlShowing( object sender, DataGridViewEditingControlShowingEventArgs e )
        {
            var txtBox = e.Control as TextBox;
            if (e.Control is TextBox)
            {
                if (txtBox != null)
                {
                    txtBox.TextChanged += new EventHandler(txtBox_TextChanged);
                    txtBox.KeyPress += new KeyPressEventHandler(txtBox_KeyPress);
                }
            }
        }

        void txtBox_TextChanged( object sender, EventArgs e )
        {
            if (( sender as TextBox ).Text == "")
            {
                ( sender as TextBox ).Text = "0";
            }
        }

        void txtBox_KeyPress( object sender, KeyPressEventArgs e )
        {
            if (!char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar))
            {
                e.Handled = true;
            }

            if (( sender as TextBox ).Text.Length == 0)
            {
                ( sender as TextBox ).Text = "0";
            }
        }
        #endregion

        private void btnLoadDonGia_Click( object sender, EventArgs e )
        {
            GetAllDonGia();
        }

        private void cbShowPass_CheckedChanged( object sender, EventArgs e )
        {
            if (cbShowPass.Checked == true)
            {
                tbPass.PasswordChar = '\0';
            }
            else
            {
                tbPass.PasswordChar = '*';
            }
        }

        private void  LoadRealSchedual( DateTime fromDate, DateTime toDate )
        {   
            string[,] realSchedualArray = busWorking.GetSchedualArray("REAL", fromDate, toDate);
            ViewToDatagrid(realSchedualArray);
        }

        private void LoadFakeSchedual( DateTime fromDate, DateTime toDate )
        {
            
            string[,] fakeSchedualArray = busWorking.GetSchedualArray("FAKE", fromDate, toDate);
            ViewToDatagrid(fakeSchedualArray);
        }

        private void ViewToDatagrid( string[,] SchedualArray )
        {
            ListSchedual.Rows.Clear();
            ListSchedual.AutoGenerateColumns = true;

            if (SchedualArray == null)
            {
                MessageBox.Show("Không tồn tại dữ liệu");
                return;
            }
            // Khai báo số cột cho Grid
            ListSchedual.ColumnCount = SchedualArray.GetLength(1);

            // Duyệt từng row
            for (int i = 0 ; i < SchedualArray.GetLength(0) ; i++)
            {
                // Tạo 1 row là 1 mảng với số cột là Length của phần tử
                string[] row = new string[SchedualArray.GetLength(1)];
                for (int j = 0 ; j < SchedualArray.GetLength(1) ; j++)
                {
                    // Gán giá trị cho row
                    row[j] = SchedualArray[i, j].ToString();
                }
                // thêm row vào datagrid
                ListSchedual.Rows.Add(row);
            }
        }      

        private void ListSchedual_CellClick( object sender, DataGridViewCellEventArgs e )
        {
            panelEditFakeData.Visible = true;   
            btnUpdateWorking.Enabled = true;
            int numrow = e.RowIndex;
            int numCol = e.ColumnIndex;
            if (numCol < 3 || numrow <2)
            {
                txtIDWorking.Text = "";
                txtHoVaTen.Text = "";
                txtMaKH.Text = "";
                txtDateWorking.Value = DateTime.Now;
                return;
            }

            string data = ListSchedual.Rows[numrow].Cells[numCol].Value.ToString();
            string[] arrayData = Regex.Split(data,"\n");
            string id = arrayData[1];
            //string id = data.Substring(data.IndexOf('\n') + 1);

            try
            {
                MT_WORKING oneDay = busWorking.GetByID(id);
                txtIDWorking.Text = oneDay.ID.ToString();
                txtHoVaTen.Text = oneDay.HO_VA_TEN;
                try
                {
                    if (!string.IsNullOrWhiteSpace(oneDay.MA_KHACH_HANG))
                    {                           
                        txtMaKH.SelectedIndex = txtMaKH.Items.IndexOf(oneDay.MA_KHACH_HANG.ToString().Trim());
                        OldMaKH = oneDay.MA_KHACH_HANG.ToString().Trim();
                    }
                    else
                    {
                        txtMaKH.SelectedIndex = -1;
                    }
                }
                catch
                {
                    MessageBox.Show("Có lỗi xảy ra với Mã Khách Hàng.");
                }

                txtDateWorking.Value = oneDay.WORKING_DAY;
            }
            catch (Exception ex)
            {
                if (ex.Message.Contains("KHONG_CO_DATA"))
                {
                    MessageBox.Show("Sai dữ liệu hoặc dữ liệu không tồn tại.");
                }
                else
                {
                    MessageBox.Show("Đã có lỗi xảy ra tại: " + ex.Message);
                }

                txtIDWorking.Text = "";
                txtHoVaTen.Text = "";
                txtMaKH.Text = "";
                txtDateWorking.Value = DateTime.Now;

            }
        }

        private void btnUpdateWorking_Click( object sender, EventArgs e )
        {
            panelEditFakeData.Visible = false;   
            btnUpdateWorking.Enabled = false;   
            MT_WORKING newWorking = busWorking.GetByID(txtIDWorking.Text);
            if (txtMaKH.SelectedItem.ToString().Equals(" ") || txtMaKH.SelectedItem.ToString().Equals(""))
            {
                newWorking.MA_KHACH_HANG = "";
            }
            else
            {
                newWorking.MA_KHACH_HANG = txtMaKH.SelectedItem.ToString().Trim();
            }
            
            bool isUpdate = false;
            try
            {
                isUpdate = busWorking.UpdateWorking(newWorking, OldMaKH);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Có lỗi xảy ra  khi cập nhật tại: "+ ex.Message);
            }
            
            if (!isUpdate)
            {
                MessageBox.Show("Không thể cập nhật!");
            }
            else
            {
                MessageBox.Show("Thành Công!");
            }
            OldMaKH = string.Empty;
            string[,] fakeSchedualArray = busWorking.GetSchedualArray("TMP", DateTime.Now, DateTime.Now);
            ViewToDatagrid(fakeSchedualArray);
        }

        private void groupBox5_Enter( object sender, EventArgs e )
        {
            panelEditFakeData.Visible = false;
        }

        private void groupBox6_Enter( object sender, EventArgs e )
        {
            panelEditFakeData.Visible = false;
        }        

        private void btn_SearchAgain_Click( object sender, EventArgs e )
        {
            string[,] fakeSchedualArray = busWorking.GetSchedualArray("TMP",DateTime.Now, DateTime.Now);
            ViewToDatagrid(fakeSchedualArray);
        }

    
        private void loadCompanyToExportData( DateTime fromDate, DateTime toDate )
        {
            try
            {
                cbCompany.DataSource = busWorking.GetListCompany(fromDate.Date, toDate.Date);
                cbCompany.DisplayMember = "KHACH_HANG" + "MA_KHACH_HANG";
                cbCompany.ValueMember = "MA_KHACH_HANG";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Có lỗi khi lấy danh sách khách hàng : " + ex.Message);
            }
        }

        private void cbFromDateExport_ValueChanged( object sender, EventArgs e )
        {
            loadCompanyToExportData(cbFromDateExport.Value, cbToDateExport.Value);
        }

        private void cbToDateExport_ValueChanged( object sender, EventArgs e )
        {
            loadCompanyToExportData(cbFromDateExport.Value, cbToDateExport.Value);
        }
    }
}
