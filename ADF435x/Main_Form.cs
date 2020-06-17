using CyUSB;
using sdpApi1;
//add by Hieu
using Syncfusion.WinForms.SmithChart;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Drawing;
// Giao tiếp qua Serial
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Numerics;
using System.Threading;
using System.Windows.Forms;
// Thêm ZedGraph
using ZedGraph;

namespace ADF435x
{
    public partial class Main_Form : Form
    {

        #region Constants

        USBDeviceList usbDevices;
        static CyFX2Device connectedDevice;
        static bool FirmwareLoaded;
        static bool XferSuccess;

        static SdpBase sdp;
        Spi session;
        Gpio g_session;

        static int buffer_length = 1;
        static byte[] buffer = new byte[5];     // (Bits per register / 8) + 1

        #endregion

        #region Global variables

        bool protocol = false;                  // false = USB adapter, true = SDP

        double RFout, REFin, RFoutMax = 4400, RFoutMin = 34.375, REFinMax = 250, PFDMax = 32, OutputChannelSpacing, INT, MOD, FRAC;

        decimal N, PFDFreq;

        int ChannelUpDownCount = 0, LoadIndex = -1;

        uint[] Reg = new uint[6];
        uint[] Rprevious = new uint[6];

        bool SweepActive = false, messageShownToUser = false;
        // Bien khi sua them vao
        int VCOSelBox;
        string DirectWriteBox;
        string ReadbackComparatorBox, ReadbackVCOBandBox, ReadbackVCOBox, ReadbackVersionBox;
        int ReadSelBox, BandSelectBox, ExtBandEnBox;
        int PDSynthBox;
        int PLLTestmodesBox;
        int SDTestmodesBox;
        int ICPADJENBox;



        // biến string để lấy dữ liệu từ Serial
        string Srl_db = String.Empty; //S11
        string Sphi_deg = String.Empty; // pha 
        string Srs = String.Empty;     // phức hóa S11, phần thực
        string Sxs = String.Empty;      // phần ảo
        string Sswr = String.Empty;    // swr, tỉ số sóng đứng
        string Sz = String.Empty;     // trở kháng 
        string Sre = String.Empty;     // phần thực để vẽ smith chart
        string Sim = String.Empty;     // phần ảo vẽ smith chart


        int status = 0; // Khai báo biến để xử lý sự kiện vẽ đồ thị
        double rl_db = 0; //convert string thành số để vẽ đồ thị
        double phi_deg = 0;
        double rs = 0;
        double xs = 0;
        double swr = 0;
        double z = 0;
        double re = 0;// dùng để vẽ smith chart
        double im = 0;// dùng để vẽ smith chart
        double frequency = 0; //lấy từ phần phát 

        Complex[] s11_open; //Calib s11 open
        Complex[] s11_short; //Calib s11 short
        Complex[] s11_load; //Calib s11 load

        int MessError = 0; //Cho việc hiển thị message

        double phi_deg_tem1;
        double phi_deg_tem2;
        const double phi_deg_range = 10.0 ; //Sai số
        const double phi_range = 10.0 * Math.PI / 180;
        //const double phi_div = 10.0; //Tỉ lệ sai số ở các điểm cực trị

        SmithLineModel model = new SmithLineModel();

        #endregion

        #region Main sections

        public Main_Form(string ADIsimPLL_import_file)
        {

            InitializeComponent();
            InitializeMenus();

            if (ADIsimPLL_import_file != "")
                importADIsimPLL(ADIsimPLL_import_file);

            usbDevices = new USBDeviceList(CyConst.DEVICES_CYUSB);

            usbDevices.DeviceAttached += new EventHandler(usbDevices_DeviceAttached);
            usbDevices.DeviceRemoved += new EventHandler(usbDevices_DeviceRemoved);

            this.FormClosing += new FormClosingEventHandler(exitEventHandler);
        }
        private void Main_Form_Load(object sender, EventArgs e)
        {
            comboBox1.DataSource = SerialPort.GetPortNames(); // Lấy nguồn cho comboBox là tên của cổng COM
            comboBox1.Text = Properties.Settings.Default.DefaultCOM; // Lấy ComName đã làm ở bước 5 cho comboBox

            // Khởi tạo ZedGraph
            GraphPane myPane = zedGraphControl1.GraphPane;
            myPane.Title.Text = "Return Loss";
            myPane.XAxis.Title.Text = "Frequency";
            myPane.YAxis.Title.Text = "Return Loss";

            RollingPointPairList list = new RollingPointPairList(60000);
            LineItem curve = myPane.AddCurve("Return Loss", list, Color.Red, SymbolType.None);
            myPane.XAxis.Scale.Min = 350;
            myPane.XAxis.Scale.Max = 2700;
            myPane.XAxis.Scale.MinorStep = 100;
            myPane.XAxis.Scale.MajorStep = 500;
            myPane.YAxis.Scale.Min = -40;
            myPane.YAxis.Scale.Max = 0;

            myPane.AxisChange();

            //S11 smith chart
            LineSeries series = new LineSeries();
            //series.ColorModel = Color.Transparent;
            series.MarkerVisible = true;
            series.MarkerHeight = 4;
            series.MarkerWidth = 4;
            series.LegendText = "S11";
            series.DataSource = model.Trace1;
            series.ResistanceMember = "Re";
            series.ReactanceMember = "Im";
            series.TooltipVisible = true;
            //sfSmithChart1.
            sfSmithChart1.Series.Add(series);
            sfSmithChart1.RadialAxis.MinorGridlinesVisible = true;
            sfSmithChart1.HorizontalAxis.MinorGridlinesVisible = true;

            sfSmithChart1.ThemeName = "Office2016White";
            sfSmithChart1.Legend.Visible = false;
        }
        // Hàm Tick này sẽ bắt sự kiện cổng Serial mở hay không
        private void timer1_Tick(object sender, EventArgs e)
        {
            if (!serialPort1.IsOpen)
            {
                return;
            }
            else if (serialPort1.IsOpen)
            {
                //progressBar1.Value = 100;
                Draw();
                Data_Listview();
                status = 0;
            }
        }

        // Hàm này lưu lại cổng COM đã chọn cho lần kết nối
        private void SaveSetting()
        {
            Properties.Settings.Default.DefaultCOM = comboBox1.Text;
            Properties.Settings.Default.Save();
        }
        //Sửa phase sau khi nhận từ arduino
        private void FixPhase(List<double> phi_deg_list)
        {
            for (int i = 0; i < phi_deg_list.Count - 2; i++)
            {
                if ((Math.Abs(phi_deg_list[i]) - Math.Abs(phi_deg_list[i + 1])) < phi_deg_range && (Math.Abs(phi_deg_list[i+1]) - Math.Abs(phi_deg_list[i + 2])) < phi_deg_range) //Đạo hàm dương
                {
                    phi_deg_list[phi_deg_list.FindIndex(ind => ind.Equals(phi_deg_list[i + 2]))] = - Math.Abs(phi_deg_list[i + 2]);
                    phi_deg_list[phi_deg_list.FindIndex(ind => ind.Equals(phi_deg_list[i + 1]))] = - Math.Abs(phi_deg_list[i + 1]);
                    phi_deg_list[phi_deg_list.FindIndex(ind => ind.Equals(phi_deg_list[i]))] = - Math.Abs(phi_deg_list[i]);
                }
                else if ((Math.Abs(phi_deg_list[i]) - Math.Abs(phi_deg_list[i + 1])) < phi_deg_range && (Math.Abs(phi_deg_list[i + 1]) - Math.Abs(phi_deg_list[i + 2])) > -phi_deg_range) //Cực đại
                {
                    phi_deg_list[phi_deg_list.FindIndex(ind => ind.Equals(phi_deg_list[i + 1]))] = -Math.Abs(phi_deg_list[i + 1]);
                    phi_deg_list[phi_deg_list.FindIndex(ind => ind.Equals(phi_deg_list[i]))] = -Math.Abs(phi_deg_list[i]);
                }
                else if ((Math.Abs(phi_deg_list[i]) - Math.Abs(phi_deg_list[i + 1])) > -phi_deg_range && (Math.Abs(phi_deg_list[i + 1]) - Math.Abs(phi_deg_list[i + 2])) > -phi_deg_range) //Cực tiểu
                {
                    phi_deg_list[phi_deg_list.FindIndex(ind => ind.Equals(phi_deg_list[i + 2]))] = -Math.Abs(phi_deg_list[i + 2]);
                }
            }
        }

        private Complex[] CalibParameterRead(String path)
        {
            string[] file_read = File.ReadAllText(path).Split(new string[] { "\n" }, StringSplitOptions.RemoveEmptyEntries); //RemoveEmptyEntries: Option này để xóa phần tử null ở cuối cùng
            Complex[] calib_para = new Complex[file_read.Length];
            for (int i = 0; i < file_read.Length; i++)
            {
                string[] temp = file_read[i].Split('+');
                calib_para[i] = Complex.FromPolarCoordinates(Math.Pow(10,double.Parse(temp[0])/20), double.Parse(temp[1]) * Math.PI / 180);
            }
            return calib_para;
        }
        // Calib số thực
        private double s11_calib_real(double s11_dut, double s11_open, double s11_short, double s11_load)
        {
            double EDF = s11_load;
            double ESF = (s11_open + s11_short - 2 * EDF) / (s11_open - s11_short);
            double ERF = -2 * (s11_open - EDF) * (s11_short - EDF) / (s11_open - s11_short);
            return (s11_dut - EDF) / (s11_dut * ESF - EDF * ESF + ERF);
        }
        //Tính toán s11 sau calib
        private Complex s11_calib(Complex s11_dut, Complex s11_open, Complex s11_short, Complex s11_load)
        {
            Complex EDF = s11_load;
            Complex ESF = (s11_open + s11_short - 2 * EDF) / (s11_open - s11_short);
            Complex ERF = -2 * (s11_open - EDF) * (s11_short - EDF) / (s11_open - s11_short);
            return (s11_dut - EDF) / (s11_dut * ESF - EDF * ESF + ERF);
        }

        // Nhận và xử lý string gửi từ Serial
        public void serialPort1_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            try
            {
                string[] measure_result = serialPort1.ReadLine().Split('|'); // Đọc một dòng của Serial, cắt chuỗi khi gặp ký tự gạch đứng
                Srl_db = measure_result[0];
                Sphi_deg = measure_result[1]; //Chờ pha khác

                double.TryParse(Srl_db, out rl_db); // Chuyển đổi sang kiểu double
                double.TryParse(Sphi_deg, out phi_deg);

                #region Fix phase
                if (phi_deg_tem1 == phi_deg_tem2 && phi_deg_tem1 == 0)
                {
                    phi_deg_tem2 = phi_deg;
                }
                else if (phi_deg_tem2 != 0 && phi_deg_tem1 == 0 && (phi_deg_tem2 - phi_deg) > -phi_deg_range)
                {
                    phi_deg_tem1 = phi_deg_tem2;
                    phi_deg_tem2 = phi_deg;
                }
                else if (phi_deg_tem2 != 0 && phi_deg_tem1 == 0 && (phi_deg_tem2 - phi_deg) < phi_deg_range)
                {
                    phi_deg_tem1 = phi_deg_tem2;
                    phi_deg_tem2 = phi_deg;
                    phi_deg = -phi_deg;
                }
                else if ((phi_deg_tem1 - phi_deg_tem2) > -phi_deg_range && (phi_deg_tem2 - phi_deg) > -phi_deg_range) //Đạo hàm âm
                {
                    phi_deg_tem1 = phi_deg_tem2;
                    phi_deg_tem2 = phi_deg;
                }
                else if ((phi_deg_tem1 - phi_deg_tem2) > -phi_deg_range && (phi_deg_tem2 - phi_deg) < phi_deg_range) //Cực tiểu
                {
                    phi_deg_tem1 = phi_deg_tem2;
                    phi_deg_tem2 = phi_deg;
                    phi_deg = -phi_deg;
                }
                else if ((phi_deg_tem1 - phi_deg_tem2) < phi_deg_range && (phi_deg_tem2 - phi_deg) < phi_deg_range) //Đạo hàm dương
                {
                    phi_deg_tem1 = phi_deg_tem2;
                    phi_deg_tem2 = phi_deg;
                    phi_deg = -phi_deg;
                }
                else if ((phi_deg_tem1 - phi_deg_tem2) < phi_deg_range && (phi_deg_tem2 - phi_deg) > -phi_deg_range) //Cực đại
                {
                    phi_deg_tem1 = phi_deg_tem2;
                    phi_deg_tem2 = phi_deg;
                }
                #endregion
                /*                //Generate phase
                                Random rand = new Random();
                                double delta_phi_slow = rand.NextDouble() * (6 - 3) + 3;
                                double delta_phi_fast = rand.NextDouble() * (9 - 6) + 6;
                                if ((frequency < 1400) || (frequency > 2600))
                                {
                                    if (phi_deg < -180 + delta_phi_slow)
                                    {
                                        phi_deg = 178;
                                    }
                                    else
                                    {
                                        phi_deg -= delta_phi_slow;
                                    }
                                }
                                else if ((frequency > 1400) & (frequency < 2600))
                                {
                                    if (phi_deg < -180 + delta_phi_fast)
                                    {
                                        phi_deg = 178;
                                    }
                                    else
                                    {
                                        phi_deg -= delta_phi_fast;
                                    }
                                }

                */
                #region Calib_Checkbox
                if (Calib.Checked)
                {
/*                    if (frequency < 1000)
                    {
                        rl_db = -1 + rand.NextDouble() * 0.5;
                    }
                    else if ((frequency > 1000) && (frequency < 1400))
                    {
                        rl_db = -2 + rand.NextDouble();
                    }
                    else if (frequency > 2600)
                    {
                        rl_db += 0.5 + rand.NextDouble();
                    } 

*/                    //Phức hóa s11_dut theo phi_deg mới và calib
                    double rl = Math.Pow(10, rl_db / 20);
                    Complex rl_cpx = new Complex(rl * Math.Cos(phi_deg * Math.PI / 180), rl * Math.Sin(phi_deg * Math.PI / 180));
                    Complex s11 = s11_calib(rl_cpx, s11_open[0], s11_short[0], s11_load[0]);
                    //Complex s11 = rl_cpx;
                    //Tính toán lại các thông số
                    rl_db = 20 * Math.Log10(s11.Magnitude);
                    phi_deg = s11.Phase * 180 / Math.PI;
                    re = s11.Real;
                    im = s11.Imaginary;

                }
                else
                {
                    double rl = Math.Pow(10, rl_db / 20);
                    re = rl * Math.Cos(phi_deg * Math.PI / 180);
                    im = rl * Math.Sin(phi_deg * Math.PI / 180);
                }
                #endregion

                //Tính trở kháng
                double denominator = (1 - re) * (1 - re) + im * im;
                rs = (1 - re * re - im * im) * 50 / denominator;
                xs = (2 * im) * 50/ denominator;

                //Hàm vẽ Smith Chart
                SmithPointModel tmp = new SmithPointModel() { Re = rs/50, Im = xs/50 };
                model.Trace1.Add(tmp);
                status = 1; // Bắt sự kiện xử lý xong chuỗi, đổi starus về 1 để hiển thị dữ liệu trong ListView và vẽ đồ thị

            }
            catch
            {
                return;
            }
        }
        // Hiển thị dữ liệu trong ListView
        private void Data_Listview()
        {
            if (status == 0)
                return;
            else
            {
                ListViewItem item = new ListViewItem(frequency.ToString()); // Gán biến frequency vào cột đầu tiên của ListView
                item.SubItems.Add(rl_db.ToString());
                item.SubItems.Add(phi_deg.ToString());
                item.SubItems.Add(rs.ToString());
                item.SubItems.Add(xs.ToString());
                item.SubItems.Add(swr.ToString());
                item.SubItems.Add(z.ToString());
                listView1.Items.Add(item); // Không nên gán string SDatas vì khi xuất dữ liệu sang Excel sẽ là dạng string, không thực hiện các phép toán được

                listView1.Items[listView1.Items.Count - 1].EnsureVisible(); // Hiện thị dòng được gán gần nhất ở ListView, tức là mình cuộn ListView theo dữ liệu gần nhất đó
            }
        }
        // Vẽ đồ thị
        private void Draw()
        {

            if (zedGraphControl1.GraphPane.CurveList.Count <= 0)
                return;

            LineItem curve = zedGraphControl1.GraphPane.CurveList[0] as LineItem;

            if (curve == null)
                return;

            IPointListEdit list = curve.Points as IPointListEdit;

            if (list == null)
                return;

            list.Add(frequency, (rl_db < 0) ? rl_db : 0); // Thêm điểm trên đồ thị

            Scale xScale = zedGraphControl1.GraphPane.XAxis.Scale;
            Scale yScale = zedGraphControl1.GraphPane.YAxis.Scale;


            //// Tự động Scale theo trục x
            //if (frequency > xScale.Max - xScale.MajorStep)
            //{
            //    xScale.Max = frequency + xScale.MajorStep;
            //    xScale.Min = xScale.Max - 30;
            //}

            //// Tự động Scale theo trục y
            //if (rl_db > yScale.Max - yScale.MajorStep)
            //{
            //    yScale.Max = rl_db + yScale.MajorStep;
            //}
            //else if (rl_db < yScale.Min + yScale.MajorStep)
            //{
            //    yScale.Min = rl_db - yScale.MajorStep;
            //}

            zedGraphControl1.AxisChange();
            zedGraphControl1.Invalidate();
            zedGraphControl1.Refresh();
        }
        // Xóa đồ thị
        private void ClearZedGraph()
        {
            zedGraphControl1.GraphPane.CurveList.Clear(); // Xóa đường
            zedGraphControl1.GraphPane.GraphObjList.Clear(); // Xóa đối tượng

            zedGraphControl1.AxisChange();
            zedGraphControl1.Invalidate();

            GraphPane myPane = zedGraphControl1.GraphPane;
            myPane.Title.Text = "Đồ thị S11";
            myPane.XAxis.Title.Text = "Frequency";
            myPane.YAxis.Title.Text = "S11";

            RollingPointPairList list = new RollingPointPairList(60000);
            LineItem curve = myPane.AddCurve("S11", list, Color.Red, SymbolType.None);

            myPane.XAxis.Scale.Min = 350;
            myPane.XAxis.Scale.Max = 2700;
            myPane.XAxis.Scale.MinorStep = 100;
            myPane.XAxis.Scale.MajorStep = 500;
            myPane.YAxis.Scale.Min = -40;
            myPane.YAxis.Scale.Max = 0;

            zedGraphControl1.AxisChange();
        }
        // Xóa Smith chart
        private void ClearSmithChart()
        {
            sfSmithChart1.Series.Clear();
            model.Trace1.Clear();

            //S11 smith chart
            LineSeries series = new LineSeries();
            //series.Interior = Color.Transparent;
            series.MarkerVisible = true;
            series.MarkerHeight = 1;
            series.MarkerWidth = 1;
            series.LegendText = "S11";
            series.DataSource = model.Trace1;
            series.ResistanceMember = "Re";
            series.ReactanceMember = "Im";
            series.TooltipVisible = true;
            sfSmithChart1.Series.Add(series);

            sfSmithChart1.RadialAxis.MinorGridlinesVisible = true;
            sfSmithChart1.HorizontalAxis.MinorGridlinesVisible = true;

            sfSmithChart1.ThemeName = "Office2016White";
            sfSmithChart1.Legend.Visible = false;
        }
        // Hàm xóa dữ liệu
        private void ResetValue()
        {
            frequency = 0;
            rl_db = 0;
            Srl_db = String.Empty;
            status = 0; // Chuyển status về 0
        }
        // Hàm lưu ListView sang Excel
        private void SaveToExcel()
        {
            Microsoft.Office.Interop.Excel.Application xla = new Microsoft.Office.Interop.Excel.Application();
            xla.Visible = true;
            Microsoft.Office.Interop.Excel.Workbook wb = xla.Workbooks.Add(Microsoft.Office.Interop.Excel.XlSheetType.xlWorksheet);
            Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;

            // đặt tên cột
            Microsoft.Office.Interop.Excel.Range rg = (Microsoft.Office.Interop.Excel.Range)ws.get_Range("A1", "B1");
            ws.Cells[1, 1] = "Tần số";
            ws.Cells[1, 2] = "Return Loss";
            ws.Cells[1, 3] = "Phase";
            ws.Cells[1, 4] = "Rs";
            ws.Cells[1, 5] = "Xs";
            ws.Cells[1, 6] = "Swr";
            ws.Cells[1, 7] = "Z";
            rg.Columns.AutoFit();

            // Lưu từ ô đầu tiên của dòng thứ 2, tức ô A2
            int i = 2;
            int j = 1;

            foreach (ListViewItem comp in listView1.Items)
            {
                ws.Cells[i, j] = comp.Text.ToString();
                foreach (ListViewItem.ListViewSubItem drv in comp.SubItems)
                {
                    ws.Cells[i, j] = drv.Text.ToString();
                    j++;
                }
                j = 1;
                i++;
            }
        }
        // Sự kiện nhấn nút btConnect
        private void btConnect_Click(object sender, EventArgs e)
        {
            if (serialPort1.IsOpen)
            {
                serialPort1.WriteLine("2"); //Gửi ký tự "2" qua Serial, tương ứng với state = 2
                serialPort1.Close();
                btConnect.Text = "Connect";
                SaveSetting(); // Lưu cổng COM vào ComName
            }
            else
            {
                try
                {
                    serialPort1.PortName = comboBox1.Text; // Lấy cổng COM
                    serialPort1.BaudRate = 9600; // Baudrate là 9600, trùng với baudrate của Arduino
                    serialPort1.Open();
                    btConnect.Text = "Disconnect";
                }
                catch
                {
                    MessageBox.Show("Không thể mở cổng " + serialPort1.PortName, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
        // Sự kiện nhấn nút btSave
        private void btSave_Click(object sender, EventArgs e)
        {
            DialogResult response;
            response = MessageBox.Show("Bạn có muốn lưu số liệu?", "Lưu", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
            if (response == DialogResult.OK)
            {
                SaveToExcel(); // Thực thi hàm lưu ListView sang Excel
            }
        }
        // Sự kiện nhấn nút btRun
        private void btRun_Click(object sender, EventArgs e)
        {
            if (serialPort1.IsOpen)
            {
                serialPort1.WriteLine("9"); //Gửi kí tự 0 cho arduino, bắt đầu quá trình đo và tính toán 
            }
            else
                MessageBox.Show("Bạn không thể đo khi chưa kết nối với thiết bị", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);

        }
        // Sự kiện nhấn nút btPause
        private void btPause_Click(object sender, EventArgs e)
        {
            if (serialPort1.IsOpen)
            {
                serialPort1.WriteLine("1"); //Gửi ký tự "1" qua Serial, dừng quá trình đo
            }
            else
                MessageBox.Show("Bạn không thể dừng khi chưa kết nối với thiết bị", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        // Sự kiện nhấn nút Clear
        private void btClear_Click(object sender, EventArgs e)
        {
            if (serialPort1.IsOpen)
            {
                DialogResult response;
                response = MessageBox.Show("Bạn có chắc muốn xóa?", "Xóa dữ liệu", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                if (response == DialogResult.OK)
                {
                    if (serialPort1.IsOpen)
                    {

                        listView1.Items.Clear(); // Xóa listview

                        //Xóa đường trong đồ thị
                        ClearZedGraph();

                        //Xóa dữ liệu trong Form
                        ResetValue();

                        //xóa đường trong smith chart
                        ClearSmithChart();
                    }
                    else
                        MessageBox.Show("Bạn không thể dừng khi chưa kết nối với thiết bị", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
                MessageBox.Show("Bạn không thể xóa khi chưa kết nối với thiết bị", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        #region Calibration

        private void Calibration_TypeOpen_button_Click_1(object sender, EventArgs e)
        {
            string path = @".\Data\Calib_Open.txt";
            if (File.Exists(path))
            {
                File.Delete(path);
                File.Create(path).Close();
            }

            //Fix_phase
            List<double> rl_list = new List<double>();
            List<double> phi_deg_list = new List<double>();

            Calibration_TypeOpen_button.Enabled = false;       // Update buttons on GUI
            Stop_Calib.Enabled = true;
            SweepActive = true;

            double SweepStop = 0, SweepCur = 0, SweepSpacing = 0, SweepStart = 0;
            int SweepDelay = 0, percent = 0, seconds_remaining = 0, minutes_remaining = 0, hours_remaining = 0;

            SweepProgress.Value = 0;                // Reset progress bar

            // Take inputs from GUI, and check for valid inputs
            try
            {
                SweepStop = double.Parse(SweepStopBox.Text);   // stop freq of measure, can be use to draw graph
            }
            catch
            {
                MessageBox.Show("Invalid Stop Frequency input!\n\nEnter numeric values only.\n\nValue reset to default.", "Invalid input", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                SweepStopBox.Text = "2700".ToString();
                SweepCur = SweepStop + 1;
            }
            if (SweepStop > 8000)
            {
                SweepStop = 8000;
                SweepStopBox.Text = 8000.ToString();
            }

            try
            {
                SweepCur = double.Parse(SweepStartBox.Text); // start freq of measure
            }
            catch
            {
                MessageBox.Show("Invalid Start Frequency input!\n\nEnter numeric values only.\n\nValue reset to default.", "Invalid input", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                SweepStartBox.Text = "350".ToString();
                SweepCur = SweepStop + 1;
            }
            if (SweepCur < 1)
            {
                SweepCur = 1;
                SweepStartBox.Text = 1.ToString();
            }

            try
            {
                SweepSpacing = double.Parse(SweepSpacingBox.Text);// spacing freq of measure, may set as default
            }
            catch
            {
                MessageBox.Show("Invalid Spacing input!\n\nEnter numeric values only.\n\nValue reset to default.", "Invalid input", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                SweepSpacingBox.Text = "10".ToString();
                SweepCur = SweepStop + 1;
            }


            try
            {
                SweepDelay = int.Parse(SweepDelayBox.Text); // delay time of each freq when measuring, can choose here
            }
            catch
            {
                MessageBox.Show("Invalid Delay input!\n\nEnter numeric values only.\n\nValue reset to default.", "Invalid input", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                SweepDelayBox.Text = "2".ToString();
                SweepCur = SweepStop + 1;
            }

            SweepStart = SweepCur;

            SweepCurrent.BackColor = Color.GreenYellow;
            SweepStopButton.BackColor = Color.GreenYellow;

            while ((SweepCur <= SweepStop) & SweepActive)
            {
                frequency = SweepCur;
                RFOutFreqBox.Text = SweepCur.ToString("0.000");
                BuildRegisters();
                //Check Serial port
                if (serialPort1.IsOpen)
                {
                    serialPort1.WriteLine("9"); //Gửi kí tự 0 cho arduino, bắt đầu quá trình đo và tính toán 
                }
                else if (MessError <= 0)
                {
                    MessageBox.Show("Bạn không thể đo khi chưa kết nối với thiết bị", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    MessError++;
                }
                else
                {
                    MessError = 0;
                    break;
                }
                WriteAllButton.PerformClick();

                SweepCurrent.Text = SweepCur.ToString("0.000");

                SweepCur += SweepSpacing;

                percent = (int)(((SweepCur - SweepStart) / (SweepStop - SweepStart)) * 100);
                percent = (percent > 100) ? 100 : percent;
                SweepProgress.Value = percent;
                SweepPercentage.Text = percent.ToString() + "%";

                #region Timer

                seconds_remaining = (int)((((SweepStop - SweepCur) / SweepSpacing) * (SweepDelay + 15)) / 1000) + 1; // 15 is for time taking for execution

                while (seconds_remaining > 59)
                {
                    minutes_remaining++;
                    seconds_remaining -= 60;
                }
                while (minutes_remaining > 59)
                {
                    hours_remaining++;
                    minutes_remaining -= 60;
                }
                if (hours_remaining > 24)
                {
                    time_remaining.Text = "1 day+";
                }
                else
                    time_remaining.Text = hours_remaining.ToString("00") + ":" + minutes_remaining.ToString("00") + ":" + seconds_remaining.ToString("00");

                seconds_remaining = minutes_remaining = hours_remaining = 0;

                #endregion

                System.Windows.Forms.Application.DoEvents();
                Thread.Sleep(SweepDelay);

                if (SweepReturnStartBox.Checked)
                {
                    if (SweepCur == SweepStop)
                        SweepCur = SweepStart;
                }
                // Tạo file calib
                rl_list.Add(Math.Pow(10, rl_db / 20));
                phi_deg_list.Add(phi_deg * Math.PI / 180);
            }
            FixPhase(phi_deg_list);
            using (StreamWriter sw = new StreamWriter(path))
            {
                for(int i = 0; i < rl_list.Count;i++)
                {
                    sw.WriteLine(rl_list[i] + "+" + phi_deg_list[i]);
                }
                sw.Close();
            }
            percent = 0;
            SweepCur = 0;
            Stop_Calib.PerformClick();
        }

        private void Calibration_TypeShort_button_Click_1(object sender, EventArgs e)
        {
            string path = @".\Data\Calib_Short.txt";
            if (File.Exists(path))
            {
                File.Delete(path);
                File.Create(path).Close();
            }

            //Fix_phase
            List<double> rl_list = new List<double>();
            List<double> phi_deg_list = new List<double>();

            Calibration_TypeShort_button.Enabled = false;       // Update buttons on GUI
            Stop_Calib.Enabled = true;
            SweepActive = true;

            double SweepStop = 0, SweepCur = 0, SweepSpacing = 0, SweepStart = 0;
            int SweepDelay = 0, percent = 0, seconds_remaining = 0, minutes_remaining = 0, hours_remaining = 0;

            SweepProgress.Value = 0;                // Reset progress bar

            // Take inputs from GUI, and check for valid inputs
            try
            {
                SweepStop = double.Parse(SweepStopBox.Text);   // stop freq of measure, can be use to draw graph
            }
            catch
            {
                MessageBox.Show("Invalid Stop Frequency input!\n\nEnter numeric values only.\n\nValue reset to default.", "Invalid input", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                SweepStopBox.Text = "2700".ToString();
                SweepCur = SweepStop + 1;
            }
            if (SweepStop > 8000)
            {
                SweepStop = 8000;
                SweepStopBox.Text = 8000.ToString();
            }

            try
            {
                SweepCur = double.Parse(SweepStartBox.Text); // start freq of measure
            }
            catch
            {
                MessageBox.Show("Invalid Start Frequency input!\n\nEnter numeric values only.\n\nValue reset to default.", "Invalid input", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                SweepStartBox.Text = "350".ToString();
                SweepCur = SweepStop + 1;
            }
            if (SweepCur < 1)
            {
                SweepCur = 1;
                SweepStartBox.Text = 1.ToString();
            }

            try
            {
                SweepSpacing = double.Parse(SweepSpacingBox.Text);// spacing freq of measure, may set as default
            }
            catch
            {
                MessageBox.Show("Invalid Spacing input!\n\nEnter numeric values only.\n\nValue reset to default.", "Invalid input", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                SweepSpacingBox.Text = "10".ToString();
                SweepCur = SweepStop + 1;
            }


            try
            {
                SweepDelay = int.Parse(SweepDelayBox.Text); // delay time of each freq when measuring, can choose here
            }
            catch
            {
                MessageBox.Show("Invalid Delay input!\n\nEnter numeric values only.\n\nValue reset to default.", "Invalid input", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                SweepDelayBox.Text = "2".ToString();
                SweepCur = SweepStop + 1;
            }

            SweepStart = SweepCur;

            SweepCurrent.BackColor = Color.GreenYellow;
            SweepStopButton.BackColor = Color.GreenYellow;

            while ((SweepCur <= SweepStop) & SweepActive)
            {
                frequency = SweepCur;
                RFOutFreqBox.Text = SweepCur.ToString("0.000");
                BuildRegisters();
                //Check Serial port
                if (serialPort1.IsOpen)
                {
                    serialPort1.WriteLine("9"); //Gửi kí tự 0 cho arduino, bắt đầu quá trình đo và tính toán 
                }
                else if (MessError <= 0)
                {
                    MessageBox.Show("Bạn không thể đo khi chưa kết nối với thiết bị", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    MessError++;
                }
                else
                {
                    MessError = 0;
                    break;
                }
                WriteAllButton.PerformClick();

                SweepCurrent.Text = SweepCur.ToString("0.000");

                SweepCur += SweepSpacing;

                percent = (int)(((SweepCur - SweepStart) / (SweepStop - SweepStart)) * 100);
                percent = (percent > 100) ? 100 : percent;
                SweepProgress.Value = percent;
                SweepPercentage.Text = percent.ToString() + "%";

                #region Timer

                seconds_remaining = (int)((((SweepStop - SweepCur) / SweepSpacing) * (SweepDelay + 15)) / 1000) + 1; // 15 is for time taking for execution

                while (seconds_remaining > 59)
                {
                    minutes_remaining++;
                    seconds_remaining -= 60;
                }
                while (minutes_remaining > 59)
                {
                    hours_remaining++;
                    minutes_remaining -= 60;
                }
                if (hours_remaining > 24)
                {
                    time_remaining.Text = "1 day+";
                }
                else
                    time_remaining.Text = hours_remaining.ToString("00") + ":" + minutes_remaining.ToString("00") + ":" + seconds_remaining.ToString("00");

                seconds_remaining = minutes_remaining = hours_remaining = 0;

                #endregion

                System.Windows.Forms.Application.DoEvents();
                Thread.Sleep(SweepDelay);

                if (SweepReturnStartBox.Checked)
                {
                    if (SweepCur == SweepStop)
                        SweepCur = SweepStart;
                }
                // Tạo file calib
                rl_list.Add(Math.Pow(10, rl_db / 20));
                phi_deg_list.Add(phi_deg * Math.PI / 180);
            }
            FixPhase(phi_deg_list);
            using (StreamWriter sw = new StreamWriter(path))
            {
                for (int i = 0; i < rl_list.Count; i++)
                {
                    sw.WriteLine(rl_list[i] + "+" + phi_deg_list[i]);
                }
                sw.Close();
            }
            percent = 0;
            SweepCur = 0;
            Stop_Calib.PerformClick();
        }

        private void Calibration_TypeLoad_button_Click_1(object sender, EventArgs e)
        {
            string path = @".\Data\Calib_Load.txt";
            if (File.Exists(path))
            {
                File.Delete(path);
                File.Create(path).Close();
            }

            //Fix_phase
            List<double> rl_list = new List<double>();
            List<double> phi_deg_list = new List<double>();

            Calibration_TypeLoad_button.Enabled = false;
            Stop_Calib.Enabled = true;
            SweepActive = true;

            double SweepStop = 0, SweepCur = 0, SweepSpacing = 0, SweepStart = 0;
            int SweepDelay = 0, percent = 0, seconds_remaining = 0, minutes_remaining = 0, hours_remaining = 0;

            SweepProgress.Value = 0;                // Reset progress bar

            // Take inputs from GUI, and check for valid inputs
            try
            {
                SweepStop = double.Parse(SweepStopBox.Text);   // stop freq of measure, can be use to draw graph
            }
            catch
            {
                MessageBox.Show("Invalid Stop Frequency input!\n\nEnter numeric values only.\n\nValue reset to default.", "Invalid input", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                SweepStopBox.Text = "2700".ToString();
                SweepCur = SweepStop + 1;
            }
            if (SweepStop > 8000)
            {
                SweepStop = 8000;
                SweepStopBox.Text = 8000.ToString();
            }

            try
            {
                SweepCur = double.Parse(SweepStartBox.Text); // start freq of measure
            }
            catch
            {
                MessageBox.Show("Invalid Start Frequency input!\n\nEnter numeric values only.\n\nValue reset to default.", "Invalid input", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                SweepStartBox.Text = "350".ToString();
                SweepCur = SweepStop + 1;
            }
            if (SweepCur < 1)
            {
                SweepCur = 1;
                SweepStartBox.Text = 1.ToString();
            }

            try
            {
                SweepSpacing = double.Parse(SweepSpacingBox.Text);// spacing freq of measure, may set as default
            }
            catch
            {
                MessageBox.Show("Invalid Spacing input!\n\nEnter numeric values only.\n\nValue reset to default.", "Invalid input", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                SweepSpacingBox.Text = "10".ToString();
                SweepCur = SweepStop + 1;
            }


            try
            {
                SweepDelay = int.Parse(SweepDelayBox.Text); // delay time of each freq when measuring, can choose here
            }
            catch
            {
                MessageBox.Show("Invalid Delay input!\n\nEnter numeric values only.\n\nValue reset to default.", "Invalid input", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                SweepDelayBox.Text = "2".ToString();
                SweepCur = SweepStop + 1;
            }

            SweepStart = SweepCur;

            SweepCurrent.BackColor = Color.GreenYellow;
            SweepStopButton.BackColor = Color.GreenYellow;

            while ((SweepCur <= SweepStop) & SweepActive)
            {
                frequency = SweepCur;
                RFOutFreqBox.Text = SweepCur.ToString("0.000");
                BuildRegisters();
                //Check Serial port
                if (serialPort1.IsOpen)
                {
                    serialPort1.WriteLine("9"); //Gửi kí tự 0 cho arduino, bắt đầu quá trình đo và tính toán 
                }
                else if (MessError <= 0)
                {
                    MessageBox.Show("Bạn không thể đo khi chưa kết nối với thiết bị", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    MessError++;
                }
                else
                {
                    MessError = 0;
                    break;
                }
                WriteAllButton.PerformClick();

                SweepCurrent.Text = SweepCur.ToString("0.000");

                SweepCur += SweepSpacing;

                percent = (int)(((SweepCur - SweepStart) / (SweepStop - SweepStart)) * 100);
                percent = (percent > 100) ? 100 : percent;
                SweepProgress.Value = percent;
                SweepPercentage.Text = percent.ToString() + "%";

                #region Timer

                seconds_remaining = (int)((((SweepStop - SweepCur) / SweepSpacing) * (SweepDelay + 15)) / 1000) + 1; // 15 is for time taking for execution

                while (seconds_remaining > 59)
                {
                    minutes_remaining++;
                    seconds_remaining -= 60;
                }
                while (minutes_remaining > 59)
                {
                    hours_remaining++;
                    minutes_remaining -= 60;
                }
                if (hours_remaining > 24)
                {
                    time_remaining.Text = "1 day+";
                }
                else
                    time_remaining.Text = hours_remaining.ToString("00") + ":" + minutes_remaining.ToString("00") + ":" + seconds_remaining.ToString("00");

                seconds_remaining = minutes_remaining = hours_remaining = 0;

                #endregion

                System.Windows.Forms.Application.DoEvents();
                Thread.Sleep(SweepDelay);

                if (SweepReturnStartBox.Checked)
                {
                    if (SweepCur == SweepStop)
                        SweepCur = SweepStart;
                }
                // Tạo file calib
                rl_list.Add(Math.Pow(10, rl_db / 20));
                phi_deg_list.Add(phi_deg * Math.PI / 180);
            }
            FixPhase(phi_deg_list);
            using (StreamWriter sw = new StreamWriter(path))
            {
                for (int i = 0; i < rl_list.Count; i++)
                {
                    sw.WriteLine(rl_list[i] + "+" + phi_deg_list[i]);
                }
                sw.Close();
            }
            percent = 0;
            SweepCur = 0;
            Stop_Calib.PerformClick();
        }

        private void Stop_Calib_Click(object sender, EventArgs e)
        {
            btPause_Click(sender, e); // bind nút stop của phần phát với nút pause của phần thu, xóa pause phần thu
            SweepActive = false;
            SweepProgress.Value = 100;
            SweepPercentage.Text = "100%";
            time_remaining.Text = "00:00:00";

            SweepCurrent.BackColor = Color.FromArgb(212, 208, 200);
            SweepStopButton.BackColor = Color.FromArgb(212, 208, 200);

            Stop_Calib.Enabled = false;
            Calibration_TypeOpen_button.Enabled = true;
            Calibration_TypeLoad_button.Enabled = true;
            Calibration_TypeShort_button.Enabled = true;
        }
        #endregion
        private void CallBuildRegisters(object sender, EventArgs e)
        {
            BuildRegisters();
        }


        private void BuildRegisters()
        {
            #region Declarations



            for (int i = 0; i < 6; i++)
                Rprevious[i] = Reg[i];

            #endregion

            #region Calculate N, INT, FRAC, MOD

            RFout = Convert.ToDouble(RFOutFreqBox.Text);
            REFin = Convert.ToDouble(RefFreqBox.Text);
            OutputChannelSpacing = Convert.ToDouble(OutputChannelSpacingBox.Text);

            PFDFreqBox.Text = (REFin * (RefDoublerBox.Checked ? 2 : 1) / (RefD2Box.Checked ? 2 : 1) / (double)RcounterBox.Value).ToString();
            PFDFreq = Convert.ToDecimal(PFDFreqBox.Text);







            #region Select divider
            if (RFout >= 2200)
                OutputDividerBox.Text = "1";
            if (RFout < 2200)
                OutputDividerBox.Text = "2";
            if (RFout < 1100)
                OutputDividerBox.Text = "4";
            if (RFout < 550)
                OutputDividerBox.Text = "8";
            if (RFout < 275)
                OutputDividerBox.Text = "16";
            if (RFout < 137.5)
                OutputDividerBox.Text = "32";
            if (RFout < 68.75)
                OutputDividerBox.Text = "64";
            #endregion


            if (FeedbackSelectBox.SelectedIndex == 1)
                N = ((decimal)RFout * Convert.ToInt16(OutputDividerBox.Text)) / PFDFreq;
            else
                N = ((decimal)RFout / PFDFreq);

            INT = (uint)N;
            MOD = (uint)(Math.Round(1000 * (PFDFreq / (decimal)OutputChannelSpacing)));
            FRAC = (uint)(Math.Round(((double)N - INT) * MOD));

            if (EnableGCD.Checked)
            {
                uint gcd = GCD((uint)MOD, (uint)FRAC);

                MOD = MOD / gcd;
                FRAC = FRAC / gcd;
            }

            if (MOD == 1)
                MOD = 2;

            INTBox.Text = INT.ToString();
            MODBox.Text = MOD.ToString();
            FRACBox.Text = FRAC.ToString();
            PFDBox.Text = PFDFreq.ToString();
            DivBox.Text = OutputDividerBox.Text;
            RFoutBox.Text = (((INT + (FRAC / MOD)) * (double)PFDFreq / Convert.ToInt16(DivBox.Text)) * ((FeedbackSelectBox.SelectedIndex == 1) ? 1 : Convert.ToInt16(DivBox.Text))).ToString();
            NvalueLabel.Text = (INT + (FRAC / MOD)).ToString();

            #region PFD max error check

            if ((PFDFreq > (decimal)PFDMax) && (BandSelectClockModeBox.SelectedIndex == 0))
                PFDWarningIcon.Visible = true;
            else if ((ADF4351.Checked) && (PFDFreq > (decimal)PFDMax) && (BandSelectClockModeBox.SelectedIndex == 1) && (FRAC != 0))
                PFDWarningIcon.Visible = true;
            else if ((ADF4351.Checked) && (PFDFreq > 90) && (BandSelectClockModeBox.SelectedIndex == 1) && (FRAC != 0))
                PFDWarningIcon.Visible = true;
            else
                PFDWarningIcon.Visible = false;

            #endregion


            #endregion

            #region Band Select Clock

            if (BandSelectClockAutosetBox.Checked)
            {
                if (BandSelectClockModeBox.SelectedIndex == 0)
                {
                    uint temp = (uint)Math.Round(8 * PFDFreq, 0);
                    if ((8 * PFDFreq - temp) > 0)
                        temp++;
                    temp = (temp > 255) ? 255 : temp;
                    BandSelectClockDividerBox.Value = (decimal)temp;
                }
                else
                {
                    uint temp = (uint)Math.Round((PFDFreq * 2), 0);
                    if ((2 * PFDFreq - temp) > 0)
                        temp++;
                    temp = (temp > 255) ? 255 : temp;
                    BandSelectClockDividerBox.Value = (decimal)temp;

                }
            }

            BandSelectClockFrequencyBox.Text = (1000 * PFDFreq / (uint)BandSelectClockDividerBox.Value).ToString("0.000");

            if (Convert.ToDouble(BandSelectClockFrequencyBox.Text) > 500)
            {
                BSCWarning1Icon.Visible = true;
                BSCWarning2Icon.Visible = false;
                BSCWarning3Icon.Visible = false;
            }
            else if ((Convert.ToDouble(BandSelectClockFrequencyBox.Text) > 125) & (BandSelectClockModeBox.SelectedIndex == 1) & (ADF4351.Checked))
            {
                BSCWarning1Icon.Visible = false;
                BSCWarning2Icon.Visible = false;
                BSCWarning3Icon.Visible = false;
            }
            else if ((Convert.ToDouble(BandSelectClockFrequencyBox.Text) > 125) & (BandSelectClockModeBox.SelectedIndex == 0) & (ADF4351.Checked))
            {
                BSCWarning1Icon.Visible = false;
                BSCWarning2Icon.Visible = true;
                BSCWarning3Icon.Visible = false;
            }
            else if ((Convert.ToDouble(BandSelectClockFrequencyBox.Text) > 125))
            {
                BSCWarning1Icon.Visible = false;
                BSCWarning2Icon.Visible = false;
                BSCWarning3Icon.Visible = true;
            }
            else
            {
                BSCWarning1Icon.Visible = false;
                BSCWarning2Icon.Visible = false;
                BSCWarning3Icon.Visible = false;
            }

            #endregion

            #region Filling in registers



            #region ADF4351
            if (ADF4351.Checked)
            {
                Reg[0] = (uint)(
                    ((int)INT & 0xFFFF) * Math.Pow(2, 15) +
                    ((int)FRAC & 0xFFF) * Math.Pow(2, 3) +
                    0);
                Reg[1] = (uint)(
                    PhaseAdjustBox.SelectedIndex * Math.Pow(2, 28) +
                    PrescalerBox.SelectedIndex * Math.Pow(2, 27) +
                    (double)PhaseValueBox.Value * Math.Pow(2, 15) +
                    ((int)MOD & 0xFFF) * Math.Pow(2, 3) +
                    1
                    );
                Reg[2] = (uint)(
                    LowNoiseSpurModeBox.SelectedIndex * Math.Pow(2, 29) +
                    MuxoutBox.SelectedIndex * Math.Pow(2, 26) +
                    (RefDoublerBox.Checked ? 1 : 0) * Math.Pow(2, 25) +
                    (RefD2Box.Checked ? 1 : 0) * Math.Pow(2, 24) +
                    (double)RcounterBox.Value * Math.Pow(2, 14) +
                    DoubleBuffBox.SelectedIndex * Math.Pow(2, 13) +
                    ChargePumpCurrentBox.SelectedIndex * Math.Pow(2, 9) +
                    LDFBox.SelectedIndex * Math.Pow(2, 8) +
                    LDPBox.SelectedIndex * Math.Pow(2, 7) +
                    PDPolarityBox.SelectedIndex * Math.Pow(2, 6) +
                    PowerdownBox.SelectedIndex * Math.Pow(2, 5) +
                    CP3StateBox.SelectedIndex * Math.Pow(2, 4) +
                    CounterResetBox.SelectedIndex * Math.Pow(2, 3) +
                    2
                    );
                Reg[3] = (uint)(
                    BandSelectClockModeBox.SelectedIndex * Math.Pow(2, 23) +
                    ABPBox.SelectedIndex * Math.Pow(2, 22) +
                    ChargeCancellationBox.SelectedIndex * Math.Pow(2, 21) +
                    CSRBox.SelectedIndex * Math.Pow(2, 18) +
                    CLKDivModeBox.SelectedIndex * Math.Pow(2, 15) +
                    (double)ClockDividerValueBox.Value * Math.Pow(2, 3) +
                    3
                    );
                Reg[4] = (uint)(
                    FeedbackSelectBox.SelectedIndex * Math.Pow(2, 23) +
                    Math.Log(Convert.ToInt16(OutputDividerBox.Text), 2) * Math.Pow(2, 20) +
                    (double)BandSelectClockDividerBox.Value * Math.Pow(2, 12) +
                    VCOPowerdownBox.SelectedIndex * Math.Pow(2, 11) +
                    MTLDBox.SelectedIndex * Math.Pow(2, 10) +
                    AuxOutputSelectBox.SelectedIndex * Math.Pow(2, 9) +
                    AuxOutputEnableBox.SelectedIndex * Math.Pow(2, 8) +
                    AuxOutputPowerBox.SelectedIndex * Math.Pow(2, 6) +
                    RFOutputEnableBox.SelectedIndex * Math.Pow(2, 5) +
                    RFOutputPowerBox.SelectedIndex * Math.Pow(2, 3) +
                    4
                    );
                Reg[5] = (uint)(
                    LDPinModeBox.SelectedIndex * Math.Pow(2, 22) +
                    ReadSelBox * Math.Pow(2, 21) +
                    ICPADJENBox * Math.Pow(2, 19) +
                    SDTestmodesBox * Math.Pow(2, 15) +
                    PLLTestmodesBox * Math.Pow(2, 11) +
                    PDSynthBox * Math.Pow(2, 10) +
                    ExtBandEnBox * Math.Pow(2, 9) +
                    BandSelectBox * Math.Pow(2, 5) +
                    VCOSelBox * Math.Pow(2, 3) +
                    5
                    );
            }
            #endregion




            R0Box.Text = String.Format("{0:X}", Reg[0]);
            R1Box.Text = String.Format("{0:X}", Reg[1]);
            R2Box.Text = String.Format("{0:X}", Reg[2]);
            R3Box.Text = String.Format("{0:X}", Reg[3]);
            R4Box.Text = String.Format("{0:X}", Reg[4]);
            R5Box.Text = String.Format("{0:X}", Reg[5]);


            if (Reg[0] != Rprevious[0])
                R0Box.BackColor = Color.LightGreen;
            if (Reg[1] != Rprevious[1])
                R1Box.BackColor = Color.LightGreen;
            if (Reg[2] != Rprevious[2])
                R2Box.BackColor = Color.LightGreen;
            if (Reg[3] != Rprevious[3])
                R3Box.BackColor = Color.LightGreen;
            if (Reg[4] != Rprevious[4])
                R4Box.BackColor = Color.LightGreen;
            if (Reg[5] != Rprevious[5])
                R5Box.BackColor = Color.LightGreen;

            #endregion

            #region Misc stuff and error check
            UpdateVCOChannelSpacing();
            UpdateVCOOutputFrequencyBox();

            if (CLKDivModeBox.SelectedIndex == 2)
            {
                TsyncLabel.Visible = true;
                TsyncLabel.Text = "Tsync = " + ((1 / (double)PFDFreq) * MOD * Convert.ToInt32(ClockDividerValueBox.Value)).ToString() + " us";
            }
            else
                TsyncLabel.Visible = false;

            if (Autowrite.Checked)
                WriteAllButton.PerformClick();

            if (MOD > 4095)
            {
                //log("MOD must be less than or equal to 4095.");
                //MODBox.BackColor = Color.Tomato;
                MODWarningIcon.Visible = true;
            }
            else
            {
                //MODBox.BackColor = SystemColors.Control;
                MODWarningIcon.Visible = false;
            }

            if (Convert.ToDouble(RFoutBox.Text) != Convert.ToDouble(RFOutFreqBox.Text))
            {
                //RFoutBox.BackColor = Color.Tomato;
                RFoutWarningIcon.Visible = true;
            }
            else
            {
                //RFoutBox.BackColor = SystemColors.Control;
                RFoutWarningIcon.Visible = false;
            }

            if ((PhaseAdjustBox.SelectedIndex == 1) && (EnableGCD.Checked))
            {
                PhaseAdjustWarningIcon.Visible = true;
            }
            else
                PhaseAdjustWarningIcon.Visible = false;

            Limit_Check();


            if (FeedbackSelectBox.SelectedIndex == 0)
                FeedbackFrequencyLabel.Text = Convert.ToDouble(RFoutBox.Text) + " MHz";
            else
                FeedbackFrequencyLabel.Text = (Convert.ToDouble(RFoutBox.Text) * (Convert.ToInt16(OutputDividerBox.Text))).ToString() + " MHz";

            if ((LowNoiseSpurModeBox.SelectedIndex == 3) && (MOD < 50))
                LowNoiseSpurModeWarningIcon.Visible = true;
            else
                LowNoiseSpurModeWarningIcon.Visible = false;

            WarningsCheck();

            #endregion
        }

        #endregion

        #region Device connections

        public void Connect_CyUSB()
        {

            // add thing to nullify SDP device if connected

            log("Attempting USB adapter board connection...");
            ConnectingLabel.Visible = true;
            FirmwareLoaded = false;
            Application.DoEvents();

            int PID = 0xB40D;
            int PID2 = 0xB403;

            connectedDevice = usbDevices[0x0456, PID] as CyFX2Device;
            if (connectedDevice != null)
                FirmwareLoaded = connectedDevice.LoadExternalRam(Application.StartupPath + "\\adf4xxx_usb_fw_2_0.hex");
            else
            {
                connectedDevice = usbDevices[0x0456, PID2] as CyFX2Device;
                if (connectedDevice != null)
                    FirmwareLoaded = connectedDevice.LoadExternalRam(Application.StartupPath + "\\adf4xxx_usb_fw_1_0.hex");
            }

            if (FirmwareLoaded)
            {
                log("Firmware loaded.");

                connectedDevice.ControlEndPt.Target = CyConst.TGT_DEVICE;
                connectedDevice.ControlEndPt.ReqType = CyConst.REQ_VENDOR;
                connectedDevice.ControlEndPt.Direction = CyConst.DIR_TO_DEVICE;
                connectedDevice.ControlEndPt.ReqCode = 0xDD;                       // DD references the function in the firmware ADF_uwave_2.hex to write to the chip
                connectedDevice.ControlEndPt.Value = 0;
                connectedDevice.ControlEndPt.Index = 0;

                DeviceConnectionStatus.Text = connectedDevice.FriendlyName + " connected.";
                DeviceConnectionStatus.ForeColor = Color.Green;

                log("USB adapter board connected.");
                ConnectDeviceButton.Enabled = false;
                protocol = false;

                #region USB Delay
                USBDelayBar.Visible = true;
                Thread.Sleep(1000);
                USBDelayBar.Value = 20;
                Application.DoEvents();
                Thread.Sleep(1000);
                USBDelayBar.Value = 40;
                Application.DoEvents();
                Thread.Sleep(1000);
                USBDelayBar.Value = 60;
                Application.DoEvents();
                Thread.Sleep(1000);
                USBDelayBar.Value = 80;
                Application.DoEvents();
                Thread.Sleep(1000);
                USBDelayBar.Value = 100;
                Application.DoEvents();
                USBDelayBar.Visible = false;
                log("USB ready.");
                #endregion
            }
            else
                log("No USB adapter board attached. Try unplugging and re-plugging the USB cable.");

            ConnectingLabel.Visible = false;

        }

        void usbDevices_DeviceAttached(object sender, EventArgs e)
        {
            USBEventArgs usbEvent = e as USBEventArgs;

            FirmwareLoaded = false;
            int PID = 0xB40D;
            int PID2 = 0xB403;

            connectedDevice = usbDevices[0x0456, PID] as CyFX2Device;
            if (connectedDevice != null)
                FirmwareLoaded = connectedDevice.LoadExternalRam(Application.StartupPath + "\\adf4xxx_usb_fw_2_0.hex");
            else
            {
                connectedDevice = usbDevices[0x0456, PID2] as CyFX2Device;
                if (connectedDevice != null)
                    FirmwareLoaded = connectedDevice.LoadExternalRam(Application.StartupPath + "\\adf4xxx_usb_fw_1_0.hex");
            }

            if (FirmwareLoaded)
            {
                log("Firmware loaded.");

                connectedDevice.ControlEndPt.Target = CyConst.TGT_DEVICE;
                connectedDevice.ControlEndPt.ReqType = CyConst.REQ_VENDOR;
                connectedDevice.ControlEndPt.Direction = CyConst.DIR_TO_DEVICE;
                connectedDevice.ControlEndPt.ReqCode = 0xDD;                       // DD references the function in the firmware ADF_uwave_2.hex to write to the chip
                connectedDevice.ControlEndPt.Value = 0;
                connectedDevice.ControlEndPt.Index = 0;

                DeviceConnectionStatus.Text = connectedDevice.FriendlyName + " connected.";
                DeviceConnectionStatus.ForeColor = Color.Green;

                log("USB adapter board connected.");
                ConnectDeviceButton.Enabled = false;
                protocol = false;

                #region USB Delay
                USBDelayBar.Visible = true;
                Thread.Sleep(1000);
                USBDelayBar.Value = 20;
                Application.DoEvents();
                Thread.Sleep(1000);
                USBDelayBar.Value = 40;
                Application.DoEvents();
                Thread.Sleep(1000);
                USBDelayBar.Value = 60;
                Application.DoEvents();
                Thread.Sleep(1000);
                USBDelayBar.Value = 80;
                Application.DoEvents();
                Thread.Sleep(1000);
                USBDelayBar.Value = 100;
                Application.DoEvents();
                USBDelayBar.Visible = false;
                log("USB ready.");
                #endregion
            }
            else
                log("No USB adapter board attached. Try unplugging and re-plugging the USB cable.");

            ConnectingLabel.Visible = false;
        }

        void usbDevices_DeviceRemoved(object sender, EventArgs e)
        {
            USBEventArgs usbEvent = e as USBEventArgs;

            log("USB device removal detected.");

            if (USBselector.Checked)
            {
                DeviceConnectionStatus.Text = usbEvent.FriendlyName + " removed.";
                DeviceConnectionStatus.ForeColor = Color.Tomato;
            }
            connectedDevice = null;
            ConnectDeviceButton.Enabled = true;
        }

        public void Connect_SDP()
        {
            String message;

            log("Attempting SDP connection...");

            try
            {

                sdp = new SdpBase();

                SdpManager.connectVisualStudioDialog("6065711100000001", "", false, out sdp);

                log("Flashing LED.");
                sdp.flashLed1();
                messageShownToUser = false;
                try
                {
                    sdp.programSdram(Application.StartupPath + "\\SDP_Blackfin_Firmware.ldr", true, true);
                    sdp.reportBootStatus(out message);

                    try
                    {
                        configNormal(sdp.ID1Connector, 0);
                    }
                    catch (Exception e)
                    {
                        if ((e is SdpApiErrEx) && (e as SdpApiErrEx).number == SdpApiErr.FunctionNotSupported)
                        {
                            if (e.Message.Substring(17) == "Use Connector A")
                            {
                                MessageBox.Show(e.Message.Substring(17));
                                messageShownToUser = true;
                                sdp.Unlock();
                                throw new Exception();
                                // Disconnect from SDP-B Rev B
                            }
                            else if (e.Message.Substring(17) == "For optimal performance ensure CLKOUT is disabled")
                            {
                                MessageBox.Show("Remove R57 from the SDP board to ensure optimum performance", "Warning!");
                                messageShownToUser = true;
                                // Ok to continue. User must have removed R57 to ensure expected performance
                            }
                            else
                                throw e;
                        }
                        else
                            throw e;
                    }
                }
                catch (Exception e)
                {
                    if ((e is SdpApiWarnEx) && (e as SdpApiWarnEx).number == SdpApiWarn.NonFatalFunctionNotSupported)
                    { }
                    else
                        throw e;
                }

                ConnectDeviceButton.Enabled = false;

                DeviceConnectionStatus.Text = "SDP board connected. Using " + sdp.ID1Connector.ToString(); ;
                DeviceConnectionStatus.ForeColor = Color.Green;
                protocol = true;
                log("SDP connected.");

                sdp.newSpi(sdp.ID1Connector, 0, 32, false, false, false, 4000000, 0, out session);

                sdp.newGpio(sdp.ID1Connector, out g_session);
                g_session.configOutput(0x1);
                g_session.bitSet(0x01);

                // id1Connector = ConnectorA on the adapter board
                // 0 = use SPI_SEL_A for LE
                // wordSize = 32
                // false = clock polarity
                // false = clock phase
                // 4,000,000 = clock frequency
                // 0 = frame frequency (irrelevant because only using 1 frame)
                // s = this is the 'session' of the connection for the SPI                

            }
            catch
            {
                sdp.Unlock();
                MessageBox.Show("SDP connection failed");
                log("SDP connection failed.");
            }
        }

        #endregion

        #region Top Menu options
        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {

            log("Exiting.");


            this.Close();

        }

        string version = "4.4.0";
        string version_date = "October 2014";

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Analog Devices ADF435x software - v" + version + " - " + version_date);

            // v4.4.0 - Added random hop feature
            // v4.3.6 - Changed band select clock divider autoset to 500 kHz for band select clock mode high.
            //        - Minor bugs and stuff
            // v4.3.5 - Improved maximum PFD frequency warning to handle high PFD frequencies in Int-N mode.
            // v4.3.4 - Bug fix.
            //          Added warning when using Phase adjust with MOD GCD. 
            //          Added option to disable event log during sweep.
        }

        private void SaveConfigurationStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog SaveConfiguration = new SaveFileDialog();
            SaveConfiguration.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            SaveConfiguration.Title = "Save a configuration file";
            SaveConfiguration.FileName = "ADF435x_settings.txt";

            #region Which part?
            if (ADF4351.Checked)
            {
                SaveConfiguration.FileName = "ADF4351_settings.txt";
            }

            #endregion

            SaveConfiguration.ShowDialog();
            System.IO.File.WriteAllText(SaveConfiguration.FileName, "");

            SaveControls(TabControl, ref SaveConfiguration);

            SaveConfiguration.Dispose();
        }

        private void LoadConfigurationStripMenuItem_Click(object sender, EventArgs e)
        {
            if (PartInUseLabel2.Text != "None")
            {
                OpenFileDialog LoadConfiguration = new OpenFileDialog();
                LoadConfiguration.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
                LoadConfiguration.ShowDialog();

                if (LoadConfiguration.FileName != "")
                {
                    string[] contents = System.IO.File.ReadAllLines(LoadConfiguration.FileName);
                    LoadControls(TabControl, contents);
                }

                LoadIndex = -1;
                LoadConfiguration.Dispose();
            }
            else
                MessageBox.Show("Select a part first.");

        }
        #endregion

        #region Other

        private void log(string message)
        {
            if (enableEventLogToolStripMenuItem.Checked)
            {
                DateTime time = DateTime.Now;
                string hour = time.Hour.ToString();
                string minute = time.Minute.ToString();
                string second = time.Second.ToString();

                if (hour.Length == 1)
                    hour = "0" + hour;
                if (minute.Length == 1)
                    minute = "0" + minute;
                if (second.Length == 1)
                    second = "0" + second;

                //EventLog.Text += "\r\n" + hour + ":" + minute + ":" + second + ": " + message;
                EventLog.AppendText("\r\n" + hour + ":" + minute + ":" + second + ": " + message);

                EventLog.Update();
                EventLog.SelectionStart = EventLog.Text.Length;
                EventLog.ScrollToCaret();
            }
        }

        private void ConnectDeviceButton_Click(object sender, EventArgs e)
        {
            USBselector.Checked = true;
            Connect_CyUSB();
        }

        private void USBadapterPicture_Click(object sender, EventArgs e)
        {
            USBselector.Checked = true;
        }

        private void exitEventHandler(object sender, System.EventArgs e)
        {
            try
            {
                if (sdp != null)
                    sdp.Unlock();
            }
            catch { }

            try
            {
                if (connectedDevice != null)
                {
                    connectedDevice.Reset();
                    connectedDevice.Dispose();
                }
            }
            catch { }
        }

        #endregion

        #region Write to device

        private void WriteToDevice(uint data)
        {
            uint[] toWrite = new uint[1];
            int x = 1;                                          // for checking the result of .writeU32()

            if (protocol)                                       // protocol: true = SDP, false = CyUSB
            {
                if (session != null)
                {
                    toWrite[0] = data;

                    //if (UseSPI_SEL_BOption.Checked)
                    //    session.slaveSelect = SpiSel.selB;
                    //else
                    //    session.slaveSelect = SpiSel.selA;

                    session.slaveSelect = UseSPI_SEL_BOption.Checked ? SpiSel.selB : SpiSel.selA;

                    configNormal(sdp.ID1Connector, 3);
                    g_session.bitClear(0x1);                    // Clear GPIO0 pin (LE)
                    x = session.writeU32(toWrite);              // Write SPI CLK and DATA (and CS)
                    g_session.bitSet(0x1);                      // Set GPIO0 pin (LE)
                    configQuiet(sdp.ID1Connector, 3);

                    if (x == 0)
                        log("0x" + String.Format("{0:X}", data) + " written to device.");
                }
                else
                    log("Writing failed.");
            }
            else
            {
                if (connectedDevice != null)
                {
                    for (int i = 0; i < 4; i++)
                        buffer[i] = (byte)(data >> (i * 8));

                    buffer[4] = 32;
                    buffer_length = 5;

                    XferSuccess = connectedDevice.ControlEndPt.XferData(ref buffer, ref buffer_length);

                    if (XferSuccess)
                        log("0x" + String.Format("{0:X}", data) + " written to device.");
                }
                else
                    log("Writing failed.");
            }
        }

        #endregion

        #region Change Part stuff

        private void resetToDefaultValuesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ChangePart(sender, e);
        }

        private void ResetAllBoxes()
        {
            ABPBox.Visible = true;
            ABPLabel.Visible = true;
            ChargeCancellationBox.Visible = true;
            ChargeCancellationLabel.Visible = true;
            BandSelectClockModeBox.Visible = true;
            BandSelectModeLabel.Visible = true;
            PhaseAdjustBox.Visible = true;
            PhaseAdjustLabel.Visible = true;

            RFoutMin = 34.375;
        }

        private void ChangePart(object sender, EventArgs e)
        {
            DeviceWarningIcon.Visible = true;
            WarningsCheck();
            ResetAllBoxes();

            PFDFreqBox.Text = "200";

            ADF4351.Checked = true;

            #region ADF4351
            if (ADF4351.Checked)
            {
                PartInUseLabel2.Text = "ADF4351";
            }
            #endregion

            InitializeMenus();
            BuildRegisters();


        }

        private void InitializeMenus()
        {
            FeedbackSelectBox.SelectedIndex = 1;
            PrescalerBox.SelectedIndex = 1;

            PhaseAdjustBox.SelectedIndex = 0;
            PDPolarityBox.SelectedIndex = 1;
            LowNoiseSpurModeBox.SelectedIndex = 0;
            MuxoutBox.SelectedIndex = 0;
            DoubleBuffBox.SelectedIndex = 0;
            ChargePumpCurrentBox.SelectedIndex = 7;
            LDFBox.SelectedIndex = 0;
            LDPBox.SelectedIndex = 0;

            PowerdownBox.SelectedIndex = 0;
            CP3StateBox.SelectedIndex = 0;
            CounterResetBox.SelectedIndex = 0;

            BandSelectClockModeBox.SelectedIndex = 0;
            ABPBox.SelectedIndex = 0;
            ChargeCancellationBox.SelectedIndex = 0;
            CSRBox.SelectedIndex = 0;
            CLKDivModeBox.SelectedIndex = 0;

            VCOPowerdownBox.SelectedIndex = 0;
            MTLDBox.SelectedIndex = 0;
            AuxOutputSelectBox.SelectedIndex = 0;
            AuxOutputEnableBox.SelectedIndex = 0;
            AuxOutputPowerBox.SelectedIndex = 0;
            RFOutputEnableBox.SelectedIndex = 1;
            RFOutputPowerBox.SelectedIndex = 3;

            LDPinModeBox.SelectedIndex = 1;

            ReadSelBox = 0;
            ICPADJENBox = 3;
            SDTestmodesBox = 0;
            PLLTestmodesBox = 0;
            PDSynthBox = 0;
            ExtBandEnBox = 0;
            BandSelectBox = 0;
            VCOSelBox = 0;

            SoftwareVersionLabel.Text = version;
        }

        #endregion

        #region Other stuff

        private void MainControlsTab_Enter(object sender, EventArgs e)
        {
            if (ADF4351.Checked)
                DeviceWarningIcon.Visible = false;
            else
                DeviceWarningIcon.Visible = true;

            WarningsCheck();
        }

        private void ChannelUpDownButton_ValueChanged(object sender, EventArgs e)
        {
            if (ChannelUpDownButton.Value > ChannelUpDownCount)
            {
                RFOutFreqBox.Text = (Convert.ToDouble(RFOutFreqBox.Text) + (Convert.ToDouble(OutputChannelSpacingBox.Text) / 1000)).ToString();
            }
            else
            {
                RFOutFreqBox.Text = (Convert.ToDouble(RFOutFreqBox.Text) - (Convert.ToDouble(OutputChannelSpacingBox.Text) / 1000)).ToString();
            }

            ChannelUpDownCount = (int)ChannelUpDownButton.Value;
        }

        private uint GCD(uint a, uint b)
        {
            if (a == 0)
                return b;
            if (b == 0)
                return a;

            if (a > b)
                return GCD(a % b, b);
            else
                return GCD(a, b % a);
        }

        private void websiteStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("http://www.analog.com/en/rfif-components/pll-synthesizersvcos/products/index.html");
        }

        private void ADIsimPLLLink_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("http://forms.analog.com/form_pages/rfcomms/adisimpll.asp");
        }

        private void EngineerZoneLink_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("http://ez.analog.com/community/rf");
        }

        private void aDF4350ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("http://www.analog.com/adf4350");
        }

        private void aDF4351ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("http://www.analog.com/adf4351");
        }

        //private void ADILogo_Click(object sender, EventArgs e)
        //{
        //    TestmodesGroup.Visible = !TestmodesGroup.Visible;
        //    TabControl.SelectedIndex = 4;
        //}

        private void UpdateVCOChannelSpacing()
        {
            if (Convert.ToDouble(RFOutFreqBox.Text) < 2200)
                VCOChannelSpacingBox.Text = (Convert.ToDouble(OutputChannelSpacingBox.Text) * Convert.ToInt16(OutputDividerBox.Text)).ToString();
            else
                VCOChannelSpacingBox.Text = OutputChannelSpacingBox.Text;
        }

        private void UpdateVCOOutputFrequencyBox()
        {
            if (Convert.ToDouble(RFOutFreqBox.Text) < 2200)
                VCOFreqBox.Text = (Convert.ToDouble(RFOutFreqBox.Text) * Convert.ToInt16(OutputDividerBox.Text)).ToString();
            else
                VCOFreqBox.Text = RFOutFreqBox.Text;
        }

        private void BandSelectClockAutosetBox_CheckedChanged(object sender, EventArgs e)
        {
            BandSelectClockDividerBox.Enabled = (BandSelectClockAutosetBox.Checked) ? false : true;
        }

        private void UseSPI_SEL_BOption_Click(object sender, EventArgs e)
        {
            UsingSPISELBLabel.Visible = UseSPI_SEL_BOption.Checked ? true : false;
        }

        private void enableEventLogToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (enableEventLogToolStripMenuItem.Checked)
            {
                log("Disabling event log. Re-enable in Tools menu.");
                enableEventLogToolStripMenuItem.Checked = false;
            }
            else
            {
                enableEventLogToolStripMenuItem.Checked = true;
                log("Re-enabled event log.");
            }
        }

        #endregion

        #region Error checking

        private void RFOutFreqBox_TextChanged(object sender, EventArgs e)
        {
            try
            {
                RFout = Convert.ToDouble(RFOutFreqBox.Text);

                if ((RFout > RFoutMax) || (RFout < RFoutMin))
                {
                    RFWarningIcon.Visible = true;
                }
                else
                {
                    RFWarningIcon.Visible = false;
                }

                BuildRegisters();
            }
            catch
            {
                RFWarningIcon.Visible = true;
            }
        }

        private void RefFreqBox_TextChanged(object sender, EventArgs e)
        {
            try
            {
                REFin = Convert.ToDouble(RefFreqBox.Text);

                if (REFin > REFinMax)
                {
                    //StatusBarLabel.Text = "Reference input frequency too high!";
                    //RefFreqBox.BackColor = Color.Tomato;
                    ReferenceFrequencyWarningIcon.Visible = true;
                }
                else
                {
                    //StatusBarLabel.Text = "";
                    //RefFreqBox.BackColor = Color.White;
                    ReferenceFrequencyWarningIcon.Visible = false;
                }

                BuildRegisters();
            }
            catch
            {
                ReferenceFrequencyWarningIcon.Visible = true;
            }
        }

        private void OutputChannelSpacingBox_TextChanged(object sender, EventArgs e)
        {
            try
            {
                OutputChannelSpacing = Convert.ToDouble(OutputChannelSpacingBox.Text);
                StatusBarLabel.Text = "";
                BuildRegisters();
            }
            catch
            {
                StatusBarLabel.Text = "Invalid channel spacing input. Please enter a numeric value.";
            }
        }

        private void Limit_Check()
        {

            if ((ADF4351.Checked) && (PrescalerBox.SelectedIndex == 0) && (RFout > 3600))
            {
                PrescalerWarning(true);
            }
            else
            {
                PrescalerWarning(false);
            }



            //if ((PrescalerBox.SelectedIndex == 1) & (N < 75))
            //    PrescalerWarning(true);
            //else
            //    PrescalerWarning(false);

            if ((N < 23) | (N > 65635))
            {
                //StatusBarLabel.Text = "Warning! N value should be between 23 and 65535 inclusive.";
                //log("Warning! N value should be between 23 and 65535 inclusive.");
                //INTBox.BackColor = Color.Tomato;
                INTWarningIcon.Visible = true;
            }
            else
            {
                //StatusBarLabel.Text = "";
                //INTBox.BackColor = SystemColors.Control;
                INTWarningIcon.Visible = false;
            }

        }

        private void PrescalerWarning(bool isError)
        {
            if (isError)
            {
                //StatusBarLabel.Text = "Warning! Prescaler output too high. Try changing prescaler...";
                //log("Warning! Prescaler output too high. Try changing prescaler...");
                //PrescalerBox.BackColor = Color.Tomato;
                PrescalerWarningIcon.Visible = true;
            }
            else
            {
                //StatusBarLabel.Text = "";
                //PrescalerBox.BackColor = Color.White;
                PrescalerWarningIcon.Visible = false;
            }
        }

        private void WarningsCheck()
        {
            if (
             (RFWarningIcon.Visible) ||
             (ReferenceFrequencyWarningIcon.Visible) ||
             (PFDWarningIcon.Visible) ||
             (PrescalerWarningIcon.Visible) ||
             (INTWarningIcon.Visible) ||
             (MODWarningIcon.Visible) ||
             (RFoutWarningIcon.Visible) ||
             (LowNoiseSpurModeWarningIcon.Visible) ||
             (BSCWarning1Icon.Visible) ||
             (BSCWarning2Icon.Visible) ||
             (BSCWarning3Icon.Visible) ||
             (PhaseAdjustWarningIcon.Visible) ||
             (DeviceWarningIcon.Visible)
             )
                WarningsPanel.Visible = true;
            else
                WarningsPanel.Visible = false;
        }

        #endregion

        #region Write Register Buttons

        private void WriteR5Button_Click(object sender, EventArgs e)
        {
            log("Writing R5...");
            WriteToDevice(Reg[5]);
            R5Box.BackColor = SystemColors.Control;
        }

        private void WriteR4Button_Click(object sender, EventArgs e)
        {
            log("Writing R4...");
            WriteToDevice(Reg[4]);
            R4Box.BackColor = SystemColors.Control;
        }

        private void WriteR3Button_Click(object sender, EventArgs e)
        {
            log("Writing R3...");
            WriteToDevice(Reg[3]);
            R3Box.BackColor = SystemColors.Control;
        }

        private void WriteR2Button_Click(object sender, EventArgs e)
        {
            log("Writing R2...");
            WriteToDevice(Reg[2]);
            R2Box.BackColor = SystemColors.Control;
        }

        private void WriteR1Button_Click(object sender, EventArgs e)
        {
            log("Writing R1...");
            WriteToDevice(Reg[1]);
            R1Box.BackColor = SystemColors.Control;
        }

        private void WriteR0Button_Click(object sender, EventArgs e)
        {
            log("Writing R0...");
            WriteToDevice(Reg[0]);
            R0Box.BackColor = SystemColors.Control;
        }

        private void WriteAllButton_Click(object sender, EventArgs e)
        {

            log("Writing R5...");
            WriteToDevice(Reg[5]);

            log("Writing R4...");
            WriteToDevice(Reg[4]);

            log("Writing R3...");
            WriteToDevice(Reg[3]);

            log("Writing R2...");
            WriteToDevice(Reg[2]);

            log("Writing R1...");
            WriteToDevice(Reg[1]);

            Thread.Sleep(10);

            log("Writing R0...");
            WriteToDevice(Reg[0]);


            R0Box.BackColor = SystemColors.Control;
            R1Box.BackColor = SystemColors.Control;
            R2Box.BackColor = SystemColors.Control;
            R3Box.BackColor = SystemColors.Control;
            R4Box.BackColor = SystemColors.Control;
            R5Box.BackColor = SystemColors.Control;
        }

        private void DirectWriteButton_Click(object sender, EventArgs e)
        {
            try
            {
                uint value = Convert.ToUInt32(DirectWriteBox, 16);
                log("Writing 0x" + String.Format("{0:X}", value));
                WriteToDevice(value);
            }
            catch { }
        }

        private void SweepStopBox_TextChanged(object sender, EventArgs e)
        {

        }

        private void zedGraphControl1_Load(object sender, EventArgs e)
        {

        }

        #endregion

        #region Registers tab

        private void FillMainControlsFromRegisters(object sender, EventArgs e)
        {
            int R0 = 0, R1 = 1, R2 = 2, R3 = 3, R4 = 4, R5 = 5;
            double i, f, m;

            StatusBarLabel.Text = "";

            #region Take inputs
            try
            {
                R0 = Convert.ToInt32(R0HexBox.Text, 16);
            }
            catch
            {
                log("Error with R0 hex input");
            }
            try
            {
                R1 = Convert.ToInt32(R1HexBox.Text, 16);
            }
            catch
            {
                log("Error with R1 hex input");
            }
            try
            {
                R2 = Convert.ToInt32(R2HexBox.Text, 16);
            }
            catch
            {
                log("Error with R2 hex input");
            }
            try
            {
                R3 = Convert.ToInt32(R3HexBox.Text, 16);
            }
            catch
            {
                log("Error with R3 hex input");
            }
            try
            {
                R4 = Convert.ToInt32(R4HexBox.Text, 16);
            }
            catch
            {
                log("Error with R4 hex input");
            }
            try
            {
                R5 = Convert.ToInt32(R5HexBox.Text, 16);
            }
            catch
            {
                log("Error with R5 hex input");
            }

            #endregion



            #region ADF4351
            if (ADF4351.Checked)
            {
                i = (R0 >> 15) & 0xFFFF;
                f = (R0 >> 3) & 0xFFF;

                PhaseAdjustBox.SelectedIndex = (R1 >> 28) & 0x1;
                PrescalerBox.SelectedIndex = (R1 >> 27) & 0x1;
                PhaseValueBox.Value = (R1 >> 15) & 0xFFF;
                m = (R1 >> 3) & 0xFFF;

                LowNoiseSpurModeBox.SelectedIndex = (R2 >> 29) & 0x3;
                MuxoutBox.SelectedIndex = (R2 >> 26) & 0x7;
                RefDoublerBox.Checked = ((((R2 >> 25) & 0x1) == 1) ? true : false);
                RefD2Box.Checked = ((((R2 >> 24) & 0x1) == 1) ? true : false);
                RcounterBox.Value = (R2 >> 14) & 0x3FF;
                DoubleBuffBox.SelectedIndex = (R2 >> 13) & 0x1;
                ChargePumpCurrentBox.SelectedIndex = (R2 >> 9) & 0xF;
                LDFBox.SelectedIndex = (R2 >> 8) & 0x1;
                LDPBox.SelectedIndex = (R2 >> 7) & 0x1;
                PDPolarityBox.SelectedIndex = (R2 >> 6) & 0x1;
                PowerdownBox.SelectedIndex = (R2 >> 5) & 0x1;
                CP3StateBox.SelectedIndex = (R2 >> 4) & 0x1;
                CounterResetBox.SelectedIndex = (R2 >> 3) & 0x1;

                BandSelectClockModeBox.SelectedIndex = (R3 >> 23) & 0x1;
                ABPBox.SelectedIndex = (R3 >> 22) & 0x1;
                ChargeCancellationBox.SelectedIndex = (R3 >> 21) & 0x1;
                CSRBox.SelectedIndex = (R3 >> 18) & 0x1;
                CLKDivModeBox.SelectedIndex = (R3 >> 15) & 0x3;
                ClockDividerValueBox.Value = (R3 >> 3) & 0xFFF;

                FeedbackSelectBox.SelectedIndex = (R4 >> 23) & 0x1;
                OutputDividerBox.Text = (Math.Pow(2, ((R4 >> 20) & 0x7))).ToString();
                BandSelectClockDividerBox.Value = (R4 >> 12) & 0xFF;
                VCOPowerdownBox.SelectedIndex = (R4 >> 11) & 0x1;
                MTLDBox.SelectedIndex = (R4 >> 10) & 0x1;
                AuxOutputSelectBox.SelectedIndex = (R4 >> 9) & 0x1;
                AuxOutputEnableBox.SelectedIndex = (R4 >> 8) & 0x1;
                AuxOutputPowerBox.SelectedIndex = (R4 >> 6) & 0x3;
                RFOutputEnableBox.SelectedIndex = (R4 >> 5) & 0x1;
                RFOutputPowerBox.SelectedIndex = (R4 >> 3) & 0x3;

                LDPinModeBox.SelectedIndex = (R5 >> 22) & 0x3;
                ReadSelBox = (R5 >> 21) & 0x1;
                ICPADJENBox = (R5 >> 19) & 0x3;
                SDTestmodesBox = (R5 >> 15) & 0xF;
                PLLTestmodesBox = (R5 >> 11) & 0xF;
                PDSynthBox = (R5 >> 10) & 0x1;
                ExtBandEnBox = (R5 >> 9) & 0x1;
                BandSelectBox = (R5 >> 5) & 0xF;
                VCOSelBox = (R5 >> 3) & 0x3;

                PFDFreq = (decimal)(Convert.ToDouble(RefFreqBox.Text) / (double)RcounterBox.Value * ((RefDoublerBox.Checked) ? 2 : 1) * ((RefD2Box.Checked) ? 0.5 : 1));
                PFDFreqBox.Text = PFDFreq.ToString();
                RFout = (double)PFDFreq * (i + (f / m)) / (Convert.ToDouble(OutputDividerBox.Text));
                RFOutFreqBox.Text = RFout.ToString();
            }
            #endregion


            Reg[0] = (uint)R0;
            Reg[1] = (uint)R1;
            Reg[2] = (uint)R2;
            Reg[3] = (uint)R3;
            Reg[4] = (uint)R4;
            Reg[5] = (uint)R5;

            R0Box.Text = String.Format("{0:X}", Reg[0]);
            R1Box.Text = String.Format("{0:X}", Reg[1]);
            R2Box.Text = String.Format("{0:X}", Reg[2]);
            R3Box.Text = String.Format("{0:X}", Reg[3]);
            R4Box.Text = String.Format("{0:X}", Reg[4]);
            R5Box.Text = String.Format("{0:X}", Reg[5]);

        }

        private void TestFillButton_Click(object sender, EventArgs e)
        {
            FillMainControlsFromRegisters(this, e);
        }

        private void SweepReturnStartBox_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void Calib_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void sfSmithChart2_Click(object sender, EventArgs e)
        {

        }

        private void chart1_Click(object sender, EventArgs e)
        {

        }

        private void RegistersTab_Enter(object sender, EventArgs e)
        {
            R0HexBox.Text = R0Box.Text;
            R1HexBox.Text = R1Box.Text;
            R2HexBox.Text = R2Box.Text;
            R3HexBox.Text = R3Box.Text;
            R4HexBox.Text = R4Box.Text;
            R5HexBox.Text = R5Box.Text;


            //if ((!ADF4351.Checked))
            //    SelectADeviceWarningLabel.Visible = true;
            //else
            //    SelectADeviceWarningLabel.Visible = false;
        }




        #endregion

        #region Sweep and hop

        private void SweepStartButton_Click(object sender, EventArgs e)
        {
            s11_open = CalibParameterRead(@".\Data\Calib_Open.txt");
            s11_load = CalibParameterRead(@".\Data\Calib_Load.txt");
            s11_short = CalibParameterRead(@".\Data\Calib_Short.txt");
            //MessageBox.Show(s11_open[0].ToString() + "+" + s11_load[0].ToString() + "+" + s11_short[0].ToString());

            SweepStartButton.Enabled = false;       // Update buttons on GUI
            SweepStopButton.Enabled = true;
            SweepActive = true;

            double SweepStop = 0, SweepCur = 0, SweepSpacing = 0, SweepStart = 0;
            int SweepDelay = 0, percent = 0, seconds_remaining = 0, minutes_remaining = 0, hours_remaining = 0;

            SweepProgress.Value = 0;                // Reset progress bar

            // Take inputs from GUI, and check for valid inputs
            try
            {
                SweepStop = double.Parse(SweepStopBox.Text);   // stop freq of measure, can be use to draw graph
            }
            catch
            {
                MessageBox.Show("Invalid Stop Frequency input!\n\nEnter numeric values only.\n\nValue reset to default.", "Invalid input", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                SweepStopBox.Text = "2700".ToString();
                SweepCur = SweepStop + 1;
            }
            if (SweepStop > 8000)
            {
                SweepStop = 8000;
                SweepStopBox.Text = 8000.ToString();
            }

            try
            {
                SweepCur = double.Parse(SweepStartBox.Text); // start freq of measure
            }
            catch
            {
                MessageBox.Show("Invalid Start Frequency input!\n\nEnter numeric values only.\n\nValue reset to default.", "Invalid input", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                SweepStartBox.Text = "350".ToString();
                SweepCur = SweepStop + 1;
            }
            if (SweepCur < 1)
            {
                SweepCur = 1;
                SweepStartBox.Text = 1.ToString();
            }

            try
            {
                SweepSpacing = double.Parse(SweepSpacingBox.Text);// spacing freq of measure, may set as default
            }
            catch
            {
                MessageBox.Show("Invalid Spacing input!\n\nEnter numeric values only.\n\nValue reset to default.", "Invalid input", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                SweepSpacingBox.Text = "10".ToString();
                SweepCur = SweepStop + 1;
            }


            try
            {
                SweepDelay = int.Parse(SweepDelayBox.Text); // delay time of each freq when measuring, can choose here
            }
            catch
            {
                MessageBox.Show("Invalid Delay input!\n\nEnter numeric values only.\n\nValue reset to default.", "Invalid input", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                SweepDelayBox.Text = "2".ToString();
                SweepCur = SweepStop + 1;
            }

            SweepStart = SweepCur;

            SweepCurrent.BackColor = Color.GreenYellow;
            SweepStopButton.BackColor = Color.GreenYellow;

            while ((SweepCur <= SweepStop) & SweepActive)
            {
                frequency = SweepCur;
                RFOutFreqBox.Text = SweepCur.ToString("0.000");
                BuildRegisters();
                //Check Serial port
                if (serialPort1.IsOpen)
                {
                    serialPort1.WriteLine("9"); //Gửi kí tự 9 cho arduino, bắt đầu quá trình đo và tính toán 
                }
                else if (MessError <= 0)
                {
                    MessageBox.Show("Bạn không thể đo khi chưa kết nối với thiết bị", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    MessError++;
                }
                else
                {
                    MessError = 0;
                    break;
                }
                WriteAllButton.PerformClick();

                SweepCurrent.Text = SweepCur.ToString("0.000");

                SweepCur += SweepSpacing;

                percent = (int)(((SweepCur - SweepStart) / (SweepStop - SweepStart)) * 100);
                percent = (percent > 100) ? 100 : percent;
                SweepProgress.Value = percent;
                SweepPercentage.Text = percent.ToString() + "%";

                #region Timer

                seconds_remaining = (int)((((SweepStop - SweepCur) / SweepSpacing) * (SweepDelay + 15)) / 1000) + 1; // 15 is for time taking for execution

                while (seconds_remaining > 59)
                {
                    minutes_remaining++;
                    seconds_remaining -= 60;
                }
                while (minutes_remaining > 59)
                {
                    hours_remaining++;
                    minutes_remaining -= 60;
                }
                if (hours_remaining > 24)
                {
                    time_remaining.Text = "1 day+";
                }
                else
                    time_remaining.Text = hours_remaining.ToString("00") + ":" + minutes_remaining.ToString("00") + ":" + seconds_remaining.ToString("00");

                seconds_remaining = minutes_remaining = hours_remaining = 0;

                #endregion

                System.Windows.Forms.Application.DoEvents();
                Thread.Sleep(SweepDelay);

                if (SweepReturnStartBox.Checked)
                {
                    if (SweepCur == SweepStop)
                        SweepCur = SweepStart;
                }
                //MessageBox.Show(s11_open[0].ToString() + "+" + s11_load[0].ToString() + "+" + s11_short[0].ToString() + "+" + rl_db.ToString() + "+" + phi_deg.ToString());
                List<Complex> s11_open_list = s11_open.ToList(); s11_open_list.RemoveAt(0); s11_open = s11_open_list.ToArray();
                List<Complex> s11_short_list = s11_short.ToList(); s11_short_list.RemoveAt(0); s11_short = s11_short_list.ToArray();
                List<Complex> s11_load_list = s11_load.ToList(); s11_load_list.RemoveAt(0); s11_load = s11_load_list.ToArray();

            }

            percent = 0;
            SweepCur = 0;
            SweepStopButton.PerformClick();
        }


        private void SweepStopButton_Click(object sender, EventArgs e)
        {
            btPause_Click(sender, e); // bind nút stop của phần phát với nút pause của phần thu, xóa pause phần thu
            SweepActive = false;
            SweepProgress.Value = 100;
            SweepPercentage.Text = "100%";
            time_remaining.Text = "00:00:00";

            SweepCurrent.BackColor = Color.FromArgb(212, 208, 200);
            SweepStopButton.BackColor = Color.FromArgb(212, 208, 200);

            SweepStopButton.Enabled = false;
            SweepStartButton.Enabled = true;
        }


        #endregion

        #region SDP stuff
        private void configQuiet(SdpConnector connector, uint quietParam)
        {
            try
            {
                sdp.configQuiet(sdp.ID1Connector, quietParam);
            }
            catch (Exception e)
            {
                if ((e is SdpApiErrEx) && (e as SdpApiErrEx).number == SdpApiErr.FunctionNotSupported)
                {
                    if (messageShownToUser == false)
                    {
                        if (e.Message.Substring(17) == "SDPS: Quiet mode not supported")
                        { }
                        else
                            // Display message to user
                            throw e;
                    }
                }
                else
                    throw e;
            }
        }

        private void configNormal(SdpConnector connector, uint quietParam)
        {
            try
            {
                sdp.configNormal(sdp.ID1Connector, quietParam);
            }
            catch (Exception e)
            {
                if ((e is SdpApiErrEx) && (e as SdpApiErrEx).number == SdpApiErr.FunctionNotSupported)
                {
                    if (messageShownToUser == false)
                    {
                        if (e.Message.Substring(17) == "SDPS: Quiet mode not supported")
                        { }
                        else
                            // Display message to user
                            throw e;
                    }
                }
                else
                    throw e;
            }
        }



        private void FlashSDPLEDButton_Click(object sender, EventArgs e)
        {
            if ((protocol) & (session != null))
                sdp.flashLed1();
        }
        #endregion

        #region Save/Load configuration files
        private void SaveControls(Control ctrl, ref SaveFileDialog savefile)
        {
            foreach (Control c in ctrl.Controls)
            {
                if ((c is TextBox) | (c is CheckBox) | (c is NumericUpDown) | (c is RadioButton) | (c is ComboBox))
                {
                    if (c is TextBox)
                        System.IO.File.AppendAllText(savefile.FileName, ((System.Windows.Forms.TextBox)(c)).Text + "\r\n");

                    if (c is CheckBox)
                        System.IO.File.AppendAllText(savefile.FileName, ((System.Windows.Forms.CheckBox)(c)).Checked + "\r\n");

                    if (c is NumericUpDown)
                        System.IO.File.AppendAllText(savefile.FileName, ((System.Windows.Forms.NumericUpDown)(c)).Value + "\r\n");

                    if (c is RadioButton)
                        System.IO.File.AppendAllText(savefile.FileName, ((System.Windows.Forms.RadioButton)(c)).Checked + "\r\n");

                    if (c is ComboBox)
                        System.IO.File.AppendAllText(savefile.FileName, ((System.Windows.Forms.ComboBox)(c)).SelectedIndex + "\r\n");
                }
                else
                {
                    if (c.Controls.Count > 0)
                    {
                        SaveControls(c, ref savefile);
                    }
                }
            }
        }

        private void LoadControls(Control ctrl, string[] contents)
        {
            foreach (Control c in ctrl.Controls)
            {
                if ((c is TextBox) | (c is CheckBox) | (c is NumericUpDown) | (c is RadioButton) | (c is ComboBox))
                {
                    LoadIndex++;

                    if (c is TextBox)
                        ((System.Windows.Forms.TextBox)(c)).Text = contents[LoadIndex];

                    if (c is CheckBox)
                        ((System.Windows.Forms.CheckBox)(c)).Checked = Convert.ToBoolean(contents[LoadIndex]);

                    if (c is NumericUpDown)
                        ((System.Windows.Forms.NumericUpDown)(c)).Value = Convert.ToDecimal(contents[LoadIndex]);

                    if (c is RadioButton)
                        ((System.Windows.Forms.RadioButton)(c)).Checked = Convert.ToBoolean(contents[LoadIndex]);

                    if (c is ComboBox)
                        ((System.Windows.Forms.ComboBox)(c)).SelectedIndex = Convert.ToInt16(contents[LoadIndex]);
                }
                else
                {
                    if (c.Controls.Count > 0)
                    {
                        LoadControls(c, contents);
                    }
                }
            }
        }
        #endregion

        #region Write Hex buttons

        private void WriteR0HexButton_Click(object sender, EventArgs e)
        {
            FillMainControlsFromRegisters(this, e);
            WriteR0Button.PerformClick();
        }

        private void WriteR1HexButton_Click(object sender, EventArgs e)
        {
            FillMainControlsFromRegisters(this, e);
            WriteR1Button.PerformClick();
        }

        private void WriteR2HexButton_Click(object sender, EventArgs e)
        {
            FillMainControlsFromRegisters(this, e);
            WriteR2Button.PerformClick();
        }

        private void WriteR3HexButton_Click(object sender, EventArgs e)
        {
            FillMainControlsFromRegisters(this, e);
            WriteR3Button.PerformClick();
        }

        private void WriteR4HexButton_Click(object sender, EventArgs e)
        {
            FillMainControlsFromRegisters(this, e);
            WriteR4Button.PerformClick();
        }

        private void WriteR5HexButton_Click(object sender, EventArgs e)
        {
            FillMainControlsFromRegisters(this, e);
            WriteR5Button.PerformClick();
        }

        #endregion

        #region Import ADIsimPLL

        void importADIsimPLL(string ADIsimPLL_import_file)
        {
            IniParser.FileIniDataParser parser = new IniParser.FileIniDataParser();
            IniParser.IniData data = parser.LoadFile(ADIsimPLL_import_file);

            try
            {
                Control[] controls = this.Controls.Find(data["ChipSettings"]["Chip"], true);
                RadioButton control = controls[0] as RadioButton;
                control.Checked = true;
            }
            catch
            {
                MessageBox.Show("Invalid input file.");
            }

            resetToDefaultValuesToolStripMenuItem.PerformClick();

            #region ADF4350
            if (ADF4351.Checked)
            {

                if (data["Specifications"]["DesignFrequency"] != null)
                {
                    //RFOutFreqBox.Text = (((Convert.ToDouble(data["Specifications"]["DesignFrequency"])) / 1000000) / Convert.ToInt16(data["ChipSettings"]["VCODiv"])).ToString();
                    RFOutFreqBox.Text = (((Convert.ToDouble(data["Specifications"]["DesignFrequency"])) / 1000000)).ToString();
                }

                if (data["Specifications"]["ChSpc"] != null)
                {
                    //OutputChannelSpacingBox.Text = (((Convert.ToDouble(data["Specifications"]["ChSpc"])) / 1000) / Convert.ToInt16(data["ChipSettings"]["VCODiv"])).ToString();
                    OutputChannelSpacingBox.Text = (((Convert.ToDouble(data["Specifications"]["ChSpc"])) / 1000)).ToString();
                }

                if (data["ChipSettings"]["RefDoubler"] != null)
                    RefDoublerBox.Checked = Convert.ToBoolean(Convert.ToInt16(data["ChipSettings"]["RefDoubler"]));

                if (data["ChipSettings"]["RefDivide2"] != null)
                    RefD2Box.Checked = Convert.ToBoolean(Convert.ToInt16(data["ChipSettings"]["RefDivide2"]));

                if (data["ChipSettings"]["RefDivider"] != null)
                    RcounterBox.Value = Convert.ToInt16(data["ChipSettings"]["RefDivider"]);

                //not sure about this one
                if (data["ChipSettings"]["VCODivOutsideLoop"] != null)
                    FeedbackSelectBox.SelectedIndex = Convert.ToInt16(data["ChipSettings"]["VCODivOutsideLoop"]);

                if (data["ChipSettings"]["Prescaler"] != null)
                    PrescalerBox.SelectedIndex = (Convert.ToInt16(data["ChipSettings"]["Prescaler"]) / 4) - 1;

                if (data["ChipSettings"]["Icp_index"] != null)
                    ChargePumpCurrentBox.SelectedIndex = Convert.ToInt16(data["ChipSettings"]["Icp_index"]);

                if (data["ChipSettings"]["Dither"] != null)
                    LowNoiseSpurModeBox.SelectedIndex = Convert.ToInt16(data["ChipSettings"]["Dither"]) * 3;

                if (data["ChipSettings"]["PD_Polarity"] != null)
                    PDPolarityBox.SelectedIndex = Convert.ToInt16(data["ChipSettings"]["PD_Polarity"]);

                if (data["ChipSettings"]["CSR"] != null)
                    CSRBox.SelectedIndex = Convert.ToInt16(data["ChipSettings"]["CSR"]);


                if (data["ChipSettings"]["LockDetect"] != null)
                {
                    string temp = data["ChipSettings"]["LockDetect"];
                    if (temp == "none")
                        MuxoutBox.SelectedIndex = 0;
                    if (temp == "analogue")
                        MuxoutBox.SelectedIndex = 5;
                    if (temp == "analogue_OD")
                        MuxoutBox.SelectedIndex = 5;
                    if (temp == "digital")
                        MuxoutBox.SelectedIndex = 6;
                }

                if (data["ChipSettings"]["RFout_A"] != null)
                    RFOutputEnableBox.SelectedIndex = Convert.ToInt16(data["ChipSettings"]["RFout_A"]);

                if (data["ChipSettings"]["RFout_B"] != null)
                    AuxOutputEnableBox.SelectedIndex = Convert.ToInt16(data["ChipSettings"]["RFout_B"]);

                if (data["Reference"]["Frequency"] != null)
                    RefFreqBox.Text = ((Convert.ToDouble(data["Reference"]["Frequency"])) / 1000000).ToString();

            }
            #endregion

            BuildRegisters();

        }

        private void importADIsimPLLToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog ImportADIsimPLLDialog = new OpenFileDialog();
            ImportADIsimPLLDialog.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            ImportADIsimPLLDialog.ShowDialog();

            if (ImportADIsimPLLDialog.FileName != "")
            {
                importADIsimPLL(ImportADIsimPLLDialog.FileName);
            }

            ImportADIsimPLLDialog.Dispose();
        }

        #endregion

        #region Readback
        private void ReadbackButton_Click(object sender, EventArgs e)
        {
            if (PLLTestmodesBox != 7)
                MessageBox.Show("Readback will fail. PLL testmodes not set to READBACK to MUXOUT.");
            else if (MuxoutBox.SelectedIndex != 7)
                MessageBox.Show("Readback will fail. Muxout not set to Testmodes.");

            for (int i = 0; i < 5; i++)
                buffer[i] = 5;

            if (connectedDevice != null)
            {
                connectedDevice.ControlEndPt.Target = CyConst.TGT_DEVICE;
                connectedDevice.ControlEndPt.ReqType = CyConst.REQ_VENDOR;
                connectedDevice.ControlEndPt.Direction = CyConst.DIR_FROM_DEVICE;
                connectedDevice.ControlEndPt.ReqCode = 0xDF;                       // DD references the function in the firmware ADF_uwave_2.hex to write to the chip
                connectedDevice.ControlEndPt.Value = 0;
                connectedDevice.ControlEndPt.Index = 0;

                buffer[4] = 32;
                buffer_length = 5;

                XferSuccess = connectedDevice.ControlEndPt.XferData(ref buffer, ref buffer_length);

                int readback_value = buffer[0] << 2;
                readback_value += buffer[1] >> 6;

                if (XferSuccess & (readback_value != 0))
                    log("Readback successful.");
                else
                    log("Readback failed. Did you write to any register before clicking Readback?");


                if (ReadSelBox == 0)
                {
                    ReadbackComparatorBox = ((readback_value >> 7) & 0x7).ToString();
                    ReadbackVCOBandBox = ((readback_value >> 3) & 0xF).ToString();
                    ReadbackVCOBox = (readback_value & 0x7).ToString();

                    ReadbackVersionBox = "-";
                }
                else
                {
                    ReadbackVersionBox = readback_value.ToString();

                    ReadbackVCOBox = "-";
                    ReadbackVCOBandBox = "-";
                    ReadbackComparatorBox = "-";
                }

                connectedDevice.ControlEndPt.ReqCode = 0xDD;
                connectedDevice.ControlEndPt.Direction = CyConst.DIR_TO_DEVICE;

            }
        }
        #endregion

    }


    public class SmithPointModel
    {
        public double Re { get; set; }

        public double Im { get; set; }
    }

    public class SmithLineModel
    {
        public SmithLineModel()
        {
            Trace1 = new ObservableCollection<SmithPointModel>();
        }
        public ObservableCollection<SmithPointModel> Trace1 { get; set; }
    }
}


