using BUS;
using DTO;
using Log;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Drawing;
using System.Data;

namespace ManageWorkExpenses
{
    public partial class Main : Form
    {
        private Logger logger;
        MT_USER_BUS busUser = new MT_USER_BUS();
        MT_CONTRACT_BUS busContract = new MT_CONTRACT_BUS();
        MT_SCHEDUAL_BUS busSchedual = new MT_SCHEDUAL_BUS();
        MT_LICH_CT_BUS busCalenda = new MT_LICH_CT_BUS();

        public Main()
        {
            InitializeComponent();

            Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            config.AppSettings.File = "App.config";

            loadAllUser();
            LoadContract();
            LoadListSchedual();
            LoadListCustomer();

            string logMode = config.AppSettings.Settings["DEBUGMODE"].Value;
            if (logMode.Equals("ON"))
            {
                debugOn.Checked = true;
                debugOff.Checked = false;
            }
            else if (logMode.Equals("OFF") || string.IsNullOrEmpty(logMode))
            {
                debugOn.Checked = false;
                debugOff.Checked = true;
                config.AppSettings.Settings["DEBUGMODE"].Value = "OFF";
            }

            logger = new Logger(Utils.LogFilePath);
            logger.log("Mo chuong trinh : Main");

        }

        private void LoadListSchedual()
        {
            List<VW_SCHEDUAL> listSchedual = new List<VW_SCHEDUAL>();

            VW_SCHEDUAL day = new VW_SCHEDUAL();
            day.HO_TEN = "Ngày / Tháng";
            day.ID = 0;
            day.MA_NHAN_VIEN = null;
            day.THANG = 0;
            day.NAM = 0;
            day.TUAN1_THU2 = "31/12/18";
            day.TUAN1_THU3 = "1/1/19";
            day.TUAN1_THU4 = "2/1/19";
            day.TUAN1_THU5 = "3/1/19";
            day.TUAN1_THU6 = "4/1/19";
            day.TUAN1_THU7 = "5/1/19";
            day.TUAN1_CN   = "6/1/19";
            day.TUAN2_THU2 = "7/1/19";
            day.TUAN2_THU3 = "8/1/19";
            day.TUAN2_THU4 = "9/1/19";
            day.TUAN2_THU5 = "10/1/19";
            day.TUAN2_THU6 = "11/1/19";
            day.TUAN2_THU7 = "12/1/19";
            day.TUAN2_CN   = "13/1/19";
            day.TUAN3_THU2 = "14/1/19";
            day.TUAN3_THU3 = "15/1/19";
            day.TUAN3_THU4 = "16/1/19";
            day.TUAN3_THU5 = "17/1/19";
            day.TUAN3_THU6 = "18/1/19";
            day.TUAN3_THU7 = "19/1/19";
            day.TUAN3_CN   = "20/1/19";
            day.TUAN4_THU2 = "21/1/19";
            day.TUAN4_THU3 = "22/1/19";
            day.TUAN4_THU4 = "23/1/19";
            day.TUAN4_THU5 = "24/1/19";
            day.TUAN4_THU6 = "25/1/19";
            day.TUAN4_THU7 = "26/1/19";
            day.TUAN4_CN   = "27/1/19";
            listSchedual.Add(day);

            VW_SCHEDUAL thu = new VW_SCHEDUAL();
            thu.HO_TEN = "HỌ VÀ TÊN";
            thu.ID = 0;
            thu.MA_NHAN_VIEN = "HỌ VÀ TÊN";
            thu.THANG = 0;
            thu.NAM = 0;
            thu.TUAN1_THU2 = "2";
            thu.TUAN1_THU3 = "3";
            thu.TUAN1_THU4 = "4";
            thu.TUAN1_THU5 = "5";
            thu.TUAN1_THU6 = "6";
            thu.TUAN1_THU7 = "7";
            thu.TUAN1_CN   = "CN";
            thu.TUAN2_THU2 = "2";
            thu.TUAN2_THU3 = "3";
            thu.TUAN2_THU4 = "4";
            thu.TUAN2_THU5 = "5";
            thu.TUAN2_THU6 = "6";
            thu.TUAN2_THU7 = "7";
            thu.TUAN2_CN   = "CN";
            thu.TUAN3_THU2 = "2";
            thu.TUAN3_THU3 = "3";
            thu.TUAN3_THU4 = "4";
            thu.TUAN3_THU5 = "5";
            thu.TUAN3_THU6 = "6";
            thu.TUAN3_THU7 = "7";
            thu.TUAN3_CN   = "CN";
            thu.TUAN4_THU2 = "2";
            thu.TUAN4_THU3 = "3";
            thu.TUAN4_THU4 = "4";
            thu.TUAN4_THU5 = "5";
            thu.TUAN4_THU6 = "6";
            thu.TUAN4_THU7 = "7";
            thu.TUAN4_CN   = "CN";

            listSchedual.Add(thu);

            List<VW_SCHEDUAL> listNew = new List<VW_SCHEDUAL>();
            listNew = busSchedual.loadSchedual();
            listSchedual.AddRange(listNew);  

            ListSchedual.DataSource = listSchedual;
            ListSchedual.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

        }

        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        private void debugOn_Click( object sender, EventArgs e )
        {
            Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            config.AppSettings.File = "App.config";
            config.AppSettings.Settings["DEBUGMODE"].Value = "ON";
            config.Save();
            ConfigurationManager.RefreshSection("appSettings");
            debugOn.Checked = true;
            debugOff.Checked = false;
            logger.log("Bắt đầu ghi log : Main");
        }

        private void debugOff_Click( object sender, EventArgs e )
        {
            Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            config.AppSettings.File = "App.config";
            config.AppSettings.Settings["DEBUGMODE"].Value = "OFF";
            config.Save();
            ConfigurationManager.RefreshSection("appSettings");
            debugOn.Checked = false;
            debugOff.Checked = true;
            logger.log("Tắt log tại : Main");
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
                MT_NHAN_VIEN user = new MT_NHAN_VIEN();
                user.MA_NHAN_VIEN = tbUserCode.Text;
                user.HO_TEN = tbName.Text;
                user.CHUC_VU = tbRegency.Text;
                user.VAI_TRO = tbRole.Text;
                busUser.SaveUser(user);
                MessageBox.Show("Thành Công");
                loadAllUser();
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
            int numrow;
            numrow = e.RowIndex;
            lblIDUser.Text = ListUser.Rows[numrow].Cells[0].Value.ToString();
            tbUserCode.Text = ListUser.Rows[numrow].Cells[1].Value.ToString();
            tbName.Text = ListUser.Rows[numrow].Cells[2].Value.ToString();
            tbRegency.Text = ListUser.Rows[numrow].Cells[3].Value.ToString();
            tbRole.Text = ListUser.Rows[numrow].Cells[4].Value.ToString();
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
                                    break;
                                }
                                // MA_NHAN_VIEN
                                staff.MA_NHAN_VIEN = xlRange.Cells[i, 1].Text.ToString();

                                // HO_TEN
                                staff.HO_TEN = xlRange.Cells[i, 2].Value2.ToString();

                                // CHUC_VU
                                staff.CHUC_VU = xlRange.Cells[i, 3].Value2.ToString();

                                // VAI_TRO
                                staff.VAI_TRO = xlRange.Cells[i, 4].Value2.ToString();

                                // PHONG_BAN
                                staff.PHONG_BAN = xlRange.Cells[i, 5].Value2.ToString();

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
                                if (string.IsNullOrEmpty(xlRange.Cells[i, 1].Text.ToString()))
                                {
                                    break;
                                }
                                contract.SO_HOP_DONG = xlRange.Cells[i, 1].Text.ToString();

                                //NGAY_HOP_DONG
                                contract.NGAY_HOP_DONG = DateTime.FromOADate(double.Parse(xlRange.Cells[i, 2].Value2.ToString())).ToString("MMMM dd, yyyy");

                                //NGAY_THANH_LY
                                contract.NGAY_THANH_LY = DateTime.FromOADate(double.Parse(xlRange.Cells[i, 3].Value2.ToString())).ToString("MMMM dd, yyyy");

                                //KHACH_HANG
                                contract.KHACH_HANG = xlRange.Cells[i, 4].Value2.ToString();

                                //MA_KHACH_HANG
                                contract.MA_KHACH_HANG = xlRange.Cells[i, 5].Value2.ToString();

                                //NHOM_KHACH_HANG
                                contract.NHOM_KHACH_HANG = xlRange.Cells[i, 6].Value2.ToString();

                                //DIA_CHI
                                contract.DIA_CHI = xlRange.Cells[i, 7].Value2.ToString();

                                //TINH
                                contract.TINH = xlRange.Cells[i, 8].Value2.ToString();

                                //GIA_TRI_HOP_DONG
                                contract.GIA_TRI_HOP_DONG = xlRange.Cells[i, 9].Value2.ToString();

                                //TONG_CHI_PHI_MUC_TOI_DA
                                contract.TONG_CHI_PHI_MUC_TOI_DA = xlRange.Cells[i, 10].Value2.ToString();

                                //CHI_PHI_THUC_DA_CHI
                                contract.CHI_PHI_THUC_DA_CHI = xlRange.Cells[i, 11].Value2.ToString();

                                //GHI_CHU
                                contract.GHI_CHU = xlRange.Cells[i, 12].Text.ToString();

                                try
                                {
                                    bool result = busContract.SaveContract(contract);
                                    messeger += (result == true) ? "Ghi Thành công HĐ số : " + contract.SO_HOP_DONG + "\n" : "Không ghi được HĐ số : " + contract.SO_HOP_DONG + " Lý do: Bản ghi bị trùng số HĐ \n";
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
                ListContract.DataSource = busContract.GetListContract();
                ListContract.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Có lỗi khi lấy danh sách HĐ tại : " + ex.Message);
                //  logger.log("Có lỗi khi lấy danh sách cán bộ tại : " + ex.Message);    
            }
        }

        private void menuConfig_Click( object sender, EventArgs e )
        {
            Config setting = new Config();
            setting.ShowDialog();
        }

        private void ListContract_CellDoubleClick( object sender, DataGridViewCellEventArgs e )
        {
            try
            {
                int numrow;
                numrow = e.RowIndex;
                idHopDong.Text = ListContract.Rows[numrow].Cells[0].Value.ToString();
                tbSoHopDong.Text = ListContract.Rows[numrow].Cells[1].Value.ToString();
                tbNgayHopDong.Text = ListContract.Rows[numrow].Cells[2].Value.ToString();
                tbNgayThanhLy.Text = ListContract.Rows[numrow].Cells[3].Value.ToString();
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
            catch{}              
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
            DataGridViewCell cell2 = ListSchedual[column-1, row];
            if (cell1.Value == null || cell2.Value == null)
            {
                return false;
            }             
            return cell1.Value.ToString() == cell2.Value.ToString();
        }
        private void ListSchedual_CellPainting( object sender, DataGridViewCellPaintingEventArgs e )
        {

            // Bôi màu 2 row đầu tiên làm tiêu đề
            ListSchedual.Rows[0].DefaultCellStyle.BackColor = Color.Gray;
            ListSchedual.Rows[1].DefaultCellStyle.BackColor = Color.Gray;  
            
            // Bỏ qua không áp dụng hiệu ứng cho 2 row đầu
            if (e.RowIndex < 2 || e.ColumnIndex < 0)
                return;

            //// Bỏ border bên phải để merger.
            //e.AdvancedBorderStyle.Right = DataGridViewAdvancedCellBorderStyle.None;

            // Bôi màu cột ngày chủ nhật và cột đầu tiên.  (Chú ý thêm cột thì phải thay đổi số cho phù hợp)
            if (e.ColumnIndex == 1 || e.ColumnIndex == 12 || e.ColumnIndex == 19 || e.ColumnIndex == 26 || e.ColumnIndex == 33)
            {
                e.CellStyle.BackColor = Color.Beige;
            }


            //// Nếu các ô có cùng giá trị thì merger với nhau
            //if (IsTheSameCellValue(e.ColumnIndex, e.RowIndex))
            //{
            //    e.AdvancedBorderStyle.Left = DataGridViewAdvancedCellBorderStyle.None;               
            //}
            //else
            //{
            //    e.AdvancedBorderStyle.Left = ListSchedual.AdvancedCellBorderStyle.Left;
            //}
        }
        /// <summary>
        /// Xóa giá trị 1 ô để merge
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ListSchedual_CellFormatting( object sender, DataGridViewCellFormattingEventArgs e )
        {

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

                            // Add to MT_LICH_CT
                            MT_LICH_CT calenda = new MT_LICH_CT();
                            calenda.THANG = cbMonth.Value.Month;
                            calenda.NAM = cbYear.Value.Year;
                            calenda.FROM_DATE = DateTime.FromOADate(Convert.ToDouble(( xlWorksheet.Cells[4, 5] as Excel.Range ).Value2));
                            //DateTime.FromOADate(double.Parse(xlRange.Cells[4, 5].Value.ToString())).ToString("MMMM dd, yyyy");
                            calenda.TO_DATE = DateTime.FromOADate(Convert.ToDouble(( xlWorksheet.Cells[4, 32] as Excel.Range ).Value2));

                            bool isSuccess = busCalenda.SaveCalenda(calenda);
                            messeger+= (isSuccess == true)? "Ghi Thành công Tháng : " + calenda.THANG + "Năm :" + calenda.NAM + "\n" : messeger += "Không lưu được tháng, Dữ liệu có thể đã tồn tại. \n";
                            // Add to schedual
                            for (int i = 8 ; i <= rowCount ; i++)
                            {
                                MT_SCHEDUAL shedual = new MT_SCHEDUAL();

                                //write the value to the console 
                                //SO_HOP_DONG
                                if (string.IsNullOrEmpty(xlRange.Cells[i, 1].Text.ToString()))
                                {
                                    break;
                                }  
                                shedual.MA_NHAN_VIEN        = xlRange.Cells[i, 3].Text.ToString();
                                shedual.THANG               = cbMonth.Value.Month;
                                shedual.NAM                 = cbYear.Value.Month;
                                shedual.TUAN1_THU2          = xlRange.Cells[i, 5].Text.ToString();
                                shedual.TUAN1_THU3          = xlRange.Cells[i, 6].Text.ToString();
                                shedual.TUAN1_THU4          = xlRange.Cells[i, 7].Text.ToString();
                                shedual.TUAN1_THU5          = xlRange.Cells[i, 8].Text.ToString();
                                shedual.TUAN1_THU6          = xlRange.Cells[i, 9].Text.ToString();
                                shedual.TUAN1_THU7          = xlRange.Cells[i, 10].Text.ToString();
                                shedual.TUAN1_CN            = xlRange.Cells[i, 11].Text.ToString();
                                shedual.TUAN2_THU2          = xlRange.Cells[i, 12].Text.ToString();
                                shedual.TUAN2_THU3          = xlRange.Cells[i, 13].Text.ToString();
                                shedual.TUAN2_THU4          = xlRange.Cells[i, 14].Text.ToString();
                                shedual.TUAN2_THU5          = xlRange.Cells[i, 15].Text.ToString();
                                shedual.TUAN2_THU6          = xlRange.Cells[i, 16].Text.ToString();
                                shedual.TUAN2_THU7          = xlRange.Cells[i, 17].Text.ToString();
                                shedual.TUAN2_CN            = xlRange.Cells[i, 18].Text.ToString();
                                shedual.TUAN3_THU2          = xlRange.Cells[i, 19].Text.ToString();
                                shedual.TUAN3_THU3          = xlRange.Cells[i, 20].Text.ToString();
                                shedual.TUAN3_THU4          = xlRange.Cells[i, 21].Text.ToString();
                                shedual.TUAN3_THU5          = xlRange.Cells[i, 22].Text.ToString();
                                shedual.TUAN3_THU6          = xlRange.Cells[i, 23].Text.ToString();
                                shedual.TUAN3_THU7          = xlRange.Cells[i, 24].Text.ToString();
                                shedual.TUAN3_CN            = xlRange.Cells[i, 25].Text.ToString();
                                shedual.TUAN4_THU2          = xlRange.Cells[i, 26].Text.ToString();
                                shedual.TUAN4_THU3          = xlRange.Cells[i, 27].Text.ToString();
                                shedual.TUAN4_THU4          = xlRange.Cells[i, 28].Text.ToString();
                                shedual.TUAN4_THU5          = xlRange.Cells[i, 29].Text.ToString();
                                shedual.TUAN4_THU6          = xlRange.Cells[i, 30].Text.ToString();
                                shedual.TUAN4_THU7          = xlRange.Cells[i, 31].Text.ToString();
                                shedual.TUAN4_CN            = xlRange.Cells[i, 32].Text.ToString();   
                                try
                                {
                                    bool result = busSchedual.SaveSchedual(shedual, cbMonth.Value.Month, cbYear.Value.Year);
                                    messeger += (result == true) ? "Ghi Thành công Nhân viên: " + shedual.MA_NHAN_VIEN : "Không ghi được Nhân viên: " + shedual.MA_NHAN_VIEN;
                                }
                                catch (Exception ex)
                                {
                                    messeger += "Lỗi ghi Nhân viên: " + shedual.MA_NHAN_VIEN + " Lý do: " + ex.Message;
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
                            LoadListSchedual();
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
                bool isUpdate  = busUser.UpdateUser(user);
                string msg = "";
                msg = ( isUpdate == true ) ? "Cập nhật Thành Công!" : "Không Cập nhật được! ";
                MessageBox.Show(msg);
                loadAllUser();
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

                DialogResult dialogResult = MessageBox.Show("Bạn có chắc muốn xóa nhân viên "+ user.HO_TEN+" có Mã nhân viên là: " + user.MA_NHAN_VIEN, "Xóa Nhân Viên", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    bool isUpdate = busUser.DelUser(user);
                    string msg = "";
                    msg = ( isUpdate == true ) ? "Xóa Thành Công!" : "Không xóa được! ";                    
                    MessageBox.Show(msg);   
                    loadAllUser();
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
                contract.ID                         = int.Parse(idHopDong.Text);
                contract.SO_HOP_DONG                = tbSoHopDong.Text; 
                contract.NGAY_HOP_DONG              = tbNgayHopDong.Text;
                contract.NGAY_THANH_LY              = tbNgayThanhLy.Text;
                contract.KHACH_HANG                 = tbKhachHang.Text;
                contract.MA_KHACH_HANG              = tbMaKhachHang.Text;
                contract.NHOM_KHACH_HANG            = tbNhomKhachHang.Text;
                contract.DIA_CHI                    = tbDiaChi.Text;
                contract.TINH                       = tbTinh.Text;
                contract.GIA_TRI_HOP_DONG           = tbGiaTriHopDong.Text;
                contract.TONG_CHI_PHI_MUC_TOI_DA    = tbTongChiPhiToiDa.Text;
                contract.CHI_PHI_THUC_DA_CHI        = tbChiPhiThucDaChi.Text;
                contract.GHI_CHU                    = tbNote.Text;


                DialogResult dialogResult = MessageBox.Show("Bạn có chắc muốn xóa Hợp đồng " + contract.SO_HOP_DONG + " của Khách hàng: " + contract.KHACH_HANG, "Xóa Hợp Đồng", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    bool isUpdate = busContract.DelContract(contract);
                    string msg = "";
                    msg = ( isUpdate == true ) ? "Xóa Thành Công!" : "Không xóa được! ";
                    MessageBox.Show(msg);
                    LoadContract();
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Có lỗi khi Xóa Hợp đồng tại: " + ex.Message); 
            }

        }

        private void btnUpdateContract_Click( object sender, EventArgs e )
        {
            if (string.IsNullOrEmpty(idHopDong.Text) || idHopDong.Text.Equals("ID_Hidden"))
            {
                MessageBox.Show("Bạn chưa chọn record nào!");
                return;
            }
            if (string.IsNullOrEmpty(tbSoHopDong.Text.Trim())       || 
                string.IsNullOrEmpty(tbNgayHopDong.Text.Trim())     ||
                string.IsNullOrEmpty(tbKhachHang.Text.Trim())       ||
                string.IsNullOrEmpty(tbMaKhachHang.Text.Trim())     ||
                string.IsNullOrEmpty(tbNhomKhachHang.Text.Trim())   ||
                string.IsNullOrEmpty(tbGiaTriHopDong.Text.Trim())   ||
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
                contract.NGAY_HOP_DONG = tbNgayHopDong.Text;
                contract.NGAY_THANH_LY = tbNgayThanhLy.Text;
                contract.KHACH_HANG = tbKhachHang.Text;
                contract.MA_KHACH_HANG = tbMaKhachHang.Text;
                contract.NHOM_KHACH_HANG = tbNhomKhachHang.Text;
                contract.DIA_CHI = tbDiaChi.Text;
                contract.TINH = tbTinh.Text;
                contract.GIA_TRI_HOP_DONG = tbGiaTriHopDong.Text;
                contract.TONG_CHI_PHI_MUC_TOI_DA = tbTongChiPhiToiDa.Text;
                contract.CHI_PHI_THUC_DA_CHI = tbChiPhiThucDaChi.Text;
                contract.GHI_CHU = tbNote.Text;

                bool isUpdate = busContract.UpdateContract(contract);
                string msg = "";
                msg = ( isUpdate == true ) ? "Cập nhật Thành Công!" : "Không Cập nhật được! ";
                MessageBox.Show(msg);
                LoadContract();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Có lỗi khi Cập nhật Hợp đồng tại: " + ex.Message);  
            }
        }

        private void btnAddContract_Click( object sender, EventArgs e )
        {
            if (string.IsNullOrEmpty(tbSoHopDong.Text.Trim())       ||
                string.IsNullOrEmpty(tbNgayHopDong.Text.Trim())     ||
                string.IsNullOrEmpty(tbKhachHang.Text.Trim())       ||
                string.IsNullOrEmpty(tbMaKhachHang.Text.Trim())     ||
                string.IsNullOrEmpty(tbNhomKhachHang.Text.Trim())   ||
                string.IsNullOrEmpty(tbGiaTriHopDong.Text.Trim())   ||
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
                contract.NGAY_HOP_DONG = tbNgayHopDong.Text;
                contract.NGAY_THANH_LY = tbNgayThanhLy.Text;
                contract.KHACH_HANG = tbKhachHang.Text;
                contract.MA_KHACH_HANG = tbMaKhachHang.Text;
                contract.NHOM_KHACH_HANG = tbNhomKhachHang.Text;
                contract.DIA_CHI = tbDiaChi.Text;
                contract.TINH = tbTinh.Text;
                contract.GIA_TRI_HOP_DONG = tbGiaTriHopDong.Text;
                contract.TONG_CHI_PHI_MUC_TOI_DA = tbTongChiPhiToiDa.Text;
                contract.CHI_PHI_THUC_DA_CHI = tbChiPhiThucDaChi.Text;
                contract.GHI_CHU = tbNote.Text;

                busContract.SaveContract(contract);
                MessageBox.Show("Thành Công");
                LoadContract();
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
                //dgvListCustomer.DataSource = busContract.GetListCustomer();
                dgvListCustomer.DataSource = busContract.GetListContract();
                this.dgvListCustomer.Columns["SO_HOP_DONG"].Visible = false;
                this.dgvListCustomer.Columns["NGAY_HOP_DONG"].Visible = false;
                this.dgvListCustomer.Columns["ID"].Visible = false;
                this.dgvListCustomer.Columns["NGAY_THANH_LY"].Visible = false;
                this.dgvListCustomer.Columns["DIA_CHI"].Visible = false;
                this.dgvListCustomer.Columns["TINH"].Visible = false;
                this.dgvListCustomer.Columns["GIA_TRI_HOP_DONG"].Visible = false;
                this.dgvListCustomer.Columns["TONG_CHI_PHI_MUC_TOI_DA"].Visible = false;
                this.dgvListCustomer.Columns["CHI_PHI_THUC_DA_CHI"].Visible = false;
                this.dgvListCustomer.Columns["GHI_CHU"].Visible = false;
                this.dgvListCustomer.Columns["NHOM_KHACH_HANG"].Visible = false;
        //dgvListCustomer.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Có lỗi khi lấy danh sách khách hàng : " + ex.Message);
                //  logger.log("Có lỗi khi lấy danh sách cán bộ tại : " + ex.Message);    
            }

        }

        private void dgvListCustomer_Click(object sender, EventArgs e)
        {
            try
            {
                int index = dgvListCustomer.CurrentRow.Index;
                string strKhachHang = dgvListCustomer.Rows[index].Cells[4].Value.ToString();
                string strMaKhachHang = dgvListCustomer.Rows[index].Cells[5].Value.ToString();
                txtNameCustomer.Text = strKhachHang +" (Mã:" + strMaKhachHang + ")";
            }
            catch
            {

            }
        }

        private void btnExportexcelKQ2_Click(object sender, EventArgs e)
        {
            try
            {
                if (String.IsNullOrEmpty(txtNameCustomer.Text) )
                {
                    MessageBox.Show("Chưa chọn đơn vị khách hàng");
                }
                else
                {
                    Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                    if (xlApp == null)
                    {
                        MessageBox.Show("Excel is not properly installed!!");
                        return;
                    }
                    Microsoft.Office.Interop.Excel.Workbooks oBooks;
                    Microsoft.Office.Interop.Excel.Sheets oSheets;
                    Microsoft.Office.Interop.Excel.Workbook oBook;
                    Microsoft.Office.Interop.Excel.Worksheet oSheet;
                    //Tạo mới một Excel WorkBook 
                    xlApp.Visible = true;
                    xlApp.DisplayAlerts = false;
                    xlApp.Application.SheetsInNewWorkbook = 1;
                    oBooks = xlApp.Workbooks;
                    oBook = (Microsoft.Office.Interop.Excel.Workbook)(xlApp.Workbooks.Add(Type.Missing));
                    oSheets = oBook.Worksheets;
                    oSheet = (Microsoft.Office.Interop.Excel.Worksheet)oSheets.get_Item(1);
                    oSheet.Name = "QĐ";

                    Microsoft.Office.Interop.Excel.Range head = oSheet.get_Range("A2", "M13");
                    head.Font.Name = "Times New Roman";
                    head.Font.Size = "12";
                    head.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    // CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM
                    Microsoft.Office.Interop.Excel.Range head1 = oSheet.get_Range("A1", "M1");
                    head1.MergeCells = true;
                    head1.Value2 = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM";
                    head1.Font.Bold = false;
                    head1.Font.Name = "Times New Roman";
                    head1.Font.Size = "12";
                    head1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    
                    //Độc lập – Tự do – Hạnh phúc
                    Microsoft.Office.Interop.Excel.Range head2 = oSheet.get_Range("A2", "M2");
                    head2.MergeCells = true;
                    head2.Value2 = "Độc lập – Tự do – Hạnh phúc";
                    head2.Font.Bold = true;
                    head2.Font.Italic = true;
                    head2.Font.Underline = true;

                    Microsoft.Office.Interop.Excel.Range head3 = oSheet.get_Range("A3", "M3");
                    head3.MergeCells = true;
                    head3.Value2 = "Hà Nội, ngày .... tháng .... năm ....";
                    head3.Font.Italic = true;
                    head3.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;

                    Microsoft.Office.Interop.Excel.Range head5 = oSheet.get_Range("A5", "M5");
                    head5.MergeCells = true;
                    head5.Value2 = "QUYẾT ĐỊNH";
                    head5.Font.Bold = false;

                    Microsoft.Office.Interop.Excel.Range head6 = oSheet.get_Range("A6", "M6");
                    head6.MergeCells = true;
                    head6.Value2 = "Về việc cử cán bộ đi công tác";
                    head6.Font.Bold = false;

                    Microsoft.Office.Interop.Excel.Range head07 = oSheet.get_Range("A7", "M7");
                    head07.MergeCells = true;
                    head07.Value2 = "GIÁM ĐỐC";
                    head07.Font.Bold = false;

                    Microsoft.Office.Interop.Excel.Range head08 = oSheet.get_Range("A8", "M8");
                    head08.MergeCells = true;
                    head08.Value2 = "Công ty THNN NVC";
                    head08.Font.Bold = false;

                    Microsoft.Office.Interop.Excel.Range head10 = oSheet.get_Range("A10", "M10");
                    head10.MergeCells = true;
                    head10.Value2 = "'- Căn cứ theo Điều lệ tổ chức và hoạt động của Công ty TNHH NVC";
                    head10.Font.Italic = true;
                    head10.Font.Name = "Times New Roman";
                    head10.Font.Size = "12";
                    head10.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;

                    Microsoft.Office.Interop.Excel.Range head11 = oSheet.get_Range("A11", "M11");
                    head11.MergeCells = true;
                    head11.Value2 = "'- Căn cứ vào hợp đồng số .... ngày ......";
                    head11.Font.Italic = true;
                    head11.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;

                    Microsoft.Office.Interop.Excel.Range head12 = oSheet.get_Range("A12", "M12");
                    head12.MergeCells = true;
                    head12.Value2 = "'- Chức năng quyền hạn của Giám đốc.";
                    head12.Font.Italic = true;
                    head12.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;

                    Microsoft.Office.Interop.Excel.Range head13 = oSheet.get_Range("A13", "M13");
                    head13.MergeCells = true;
                    head13.Value2 = "'- Điều 1: Quyết định cử các nhân viên sau đi công tác:";
                    head13.Font.Italic = true;
                    head13.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                }
            }
            catch
            {

            }
        }

        private void btnExportexcelBangKe_Click(object sender, EventArgs e)
        {
           
                if (String.IsNullOrEmpty(txtNameCustomer.Text))
                {
                    MessageBox.Show("Chưa chọn đơn vị khách hàng");
                }
                else
                {
                    Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                    if (xlApp == null)
                    {
                        MessageBox.Show("Excel is not properly installed!!");
                        return;
                    }
                    Microsoft.Office.Interop.Excel.Workbooks oBooks;
                    Microsoft.Office.Interop.Excel.Sheets oSheets;
                    Microsoft.Office.Interop.Excel.Workbook oBook;
                    Microsoft.Office.Interop.Excel.Worksheet oSheet;
                    //Tạo mới một Excel WorkBook 
                    xlApp.Visible = true;
                    xlApp.DisplayAlerts = false;
                    xlApp.Application.SheetsInNewWorkbook = 1;
                    oBooks = xlApp.Workbooks;
                    oBook = (Microsoft.Office.Interop.Excel.Workbook)(xlApp.Workbooks.Add(Type.Missing));
                    oSheets = oBook.Worksheets;
                    oSheet = (Microsoft.Office.Interop.Excel.Worksheet)oSheets.get_Item(1);
                    oSheet.Name = "Bảng kê thanh toán";

                    Microsoft.Office.Interop.Excel.Range head = oSheet.get_Range("A1", "H6");
                    head.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    head.VerticalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    head.Font.Name= "Times New Roman";
                    head.Font.Bold = true;

                    Microsoft.Office.Interop.Excel.Range head1 = oSheet.get_Range("A2", "I2");
                    head1.MergeCells = true;
                    head1.Value2 = "BẢNG KÊ THANH TOÁN CÔNG TÁC PHÍ";
                    head1.Font.Size = "15";
                    Microsoft.Office.Interop.Excel.Range head2 = oSheet.get_Range("A5", "A6");
                    head2.MergeCells = true;
                    head2.Value2 = "STT";
                    head2.Font.Size = "12";
                  
                    Microsoft.Office.Interop.Excel.Range head3 = oSheet.get_Range("B5", "D6");
                    head3.MergeCells = true;
                    head3.Value2 = "Nội dung";
                    head3.Font.Size = "12";

                    Microsoft.Office.Interop.Excel.Range head4 = oSheet.get_Range("E5", "E6");
                    head4.MergeCells = true;
                    head4.Value2 = "Số ngày làm việc tại KH";
                    head4.WrapText = true;
                    head4.Font.Size = "12";

                    Microsoft.Office.Interop.Excel.Range head5 = oSheet.get_Range("F5", "F6");
                    head5.MergeCells = true;
                    head5.Value2 = "Đơn giá thanh toán";
                    head5.WrapText = true;
                    head5.Font.Size = "12";

                    Microsoft.Office.Interop.Excel.Range head6 = oSheet.get_Range("G5", "G6");
                    head6.MergeCells = true;
                    head6.Value2 = "Thành tiền";
                    head6.WrapText = true;
                    head6.Font.Size = "12";

                    Microsoft.Office.Interop.Excel.Range head7 = oSheet.get_Range("H5", "H6");
                    head7.MergeCells = true;
                    head7.Value2 = "Notes";
                    head7.Font.Size = "12";
                
            }

        }
    }
    
}
