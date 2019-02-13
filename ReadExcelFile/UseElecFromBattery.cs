using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;

namespace ReadExcelFile
{
    public partial class UseElecFromBattery : Form
    {
        List<Excel> fromExcel = new List<Excel>();
        public UseElecFromBattery()
        {
            InitializeComponent();
            dataGridView1.Visible = false;
            chart1.Visible = false;
            panel1.Visible = false;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                _FirstLoadForm();
            }
            catch (Exception ex)
            {
                MessageBox.Show("เกิดข้อผิดพลาด : " + ex.Message, "เกิดข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {

                int i = 0;
                textBox1.Clear();
                textBox2.Clear();
    
                dataGridView1.Visible = false;
                panel1.Visible = true;
                chart1.Visible = true;
                chart1.ChartAreas[0].AxisX.Title = "เวลา";
                chart1.ChartAreas[0].AxisY.Title = "แรงดัน / กระแส";
                //string[] monthNames = { "100", "75", "50", "25", "0" };
                //int startOffset = -2;
                //int endOffset = 2;
                //foreach (string monthName in monthNames)
                //{
                //    CustomLabel monthLabel = new CustomLabel(startOffset, endOffset, monthName, 0, LabelMarkStyle.None);
                //    chart1.ChartAreas[0].AxisX.CustomLabels.Add(monthLabel);
                //    startOffset = startOffset + 25;
                //    endOffset = endOffset + 25;
                //}
                //chart1.ChartAreas[0].AxisX.IntervalAutoMode = IntervalAutoMode.FixedCount;
                ////chart1.ChartAreas[0].AxisX.Maximum = fromExcel.Count();
                ////chart1.ChartAreas[0].AxisX.Minimum = 0;
                ////chart1.ChartAreas[0].AxisX.IntervalAutoMode = IntervalAutoMode.VariableCount;
                //chart1.ChartAreas[0].AxisX.IntervalType = DateTimeIntervalType.Days;
                //chart1.ChartAreas[0].AxisX.Interval = 360;

                //chart1.ChartAreas[0].AxisX.Interval = 2;

                Axis xaxis = chart1.ChartAreas[0].AxisX;
                xaxis.IntervalType = DateTimeIntervalType.Hours;
                //var total_time = (Convert.ToDateTime(fromExcel.LastOrDefault().Time) - Convert.ToDateTime(fromExcel.FirstOrDefault().Time)).TotalHours;
                //int interval = (total_time >= 9 && total_time <= 12 ? 1440
                //    : (total_time >= 6 && total_time <= 9 ? 1060
                //    : (total_time >= 3 && total_time <= 6 ? 680
                //    : (total_time >= 1 && total_time <= 3 ? 300
                //    : 1440))));
                xaxis.Interval = 1440;
                //for (int j = 1; j <= interval;j+=(interval/12) )
                //    xaxis.CustomLabels.Add(0, j, fromExcel.FirstOrDefault().Time);
                //xaxis.CustomLabels.Add(0, 25, fromExcel.FirstOrDefault().Time);
                //xaxis.CustomLabels.Add(0, 50, fromExcel.FirstOrDefault().Time);
                //xaxis.CustomLabels.Add(0, 75, fromExcel.FirstOrDefault().Time);
                
               // xaxis.IsLabelAutoFit = true;
                //xaxis.IntervalOffset = 1;
                //for (int index = 0; index < interval; index++)
                //Loop Remove point in chart
                foreach (var series in chart1.Series)
                {
                    series.Points.Clear();
                }

                //Loop Add point in chart
                foreach (var data in fromExcel)
                {
                    //string dateTime = data.Date.Replace('.','/') +  data.Time;
                    //DateTime oDate = DateTime.ParseExact(dateTime, "dd/MM/yyyy HH:mm:ss tt", null);
                    //DateTime date = Convert.ToDateTime(data.Date);
                    //Gen Graph Series Volt
                    chart1.Series["Volt"].Points.AddXY(data.Time, data.UseV);
                    chart1.Series["Volt"].Points[i].MarkerStyle = MarkerStyle.Circle;
                    chart1.Series["Volt"].Points[i].MarkerSize = (data.UseV.Equals(data.UseA) ? 5 : 3);
                    chart1.Series["Volt"].Points[i].MarkerColor = Color.Blue;
                    ////Gen Graph Series Am
                    chart1.Series["Am"].Points.AddXY(data.Time, data.UseA);
                    chart1.Series["Am"].Points[i].MarkerStyle = MarkerStyle.Circle;
                    chart1.Series["Am"].Points[i].MarkerSize = 3;
                    chart1.Series["Am"].Points[i].MarkerColor = Color.Orange;

                    
                    //chart1.Series["Am"].
                    i++;
                }
                string ArgVolt = (fromExcel.Sum(a => Convert.ToDouble(a.UseV)) / fromExcel.Count()).ToString("#.##");
                string ArgAm = (fromExcel.Sum(a => Convert.ToDouble(a.UseA)) / fromExcel.Count()).ToString("#.##");
                textBox1.AppendText(ArgVolt == "" ? "0.00" : ArgVolt);
                textBox2.AppendText(ArgAm == "" ? "0.00" : ArgAm);
            }
            catch (Exception ex)
            {
                MessageBox.Show("เกิดข้อผิดพลาด : " + ex.Message, "เกิดข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void UseElecFromBattery_Load(object sender, EventArgs e)
        {
            try
            {
                _FirstLoadForm();
            }
            catch (Exception ex)
            {
                MessageBox.Show("เกิดข้อผิดพลาด : " + ex.Message, "เกิดข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        public void _FirstLoadForm()
        {
            dataGridView1.Visible = true;
            chart1.Visible = false;
            panel1.Visible = true;
            textBox1.Clear();
            textBox2.Clear();
            ////Initialize dataGridView1
            //dataGridView1.Columns[0].HeaderText = "เวลา";
            //dataGridView1.Columns[1].HeaderText = "การใช้กระแส";
            //dataGridView1.Columns[2].HeaderText = "การใช้แรงดัน";
            // Convert to DataTable.
            DataTable table = ConvertListToDataTable(fromExcel);
            dataGridView1.DataSource = table;
            dataGridView1.Columns[1].Width = 120;
            dataGridView1.Columns[2].Width = 120;

            string ArgVolt = (fromExcel.Sum(a => Convert.ToDouble(a.UseV)) / fromExcel.Count()).ToString("#.##");
            string ArgAm = (fromExcel.Sum(a => Convert.ToDouble(a.UseA)) / fromExcel.Count()).ToString("#.##");
            textBox1.AppendText(ArgVolt == "" ? "0.00" : ArgVolt);
            textBox2.AppendText(ArgAm == "" ? "0.00" : ArgAm);
        }

        public void recieve(List<Excel> excel)
        {
            foreach (var data in excel)
            {
                fromExcel.Add(new Excel
                {
                    Date = data.Date,
                    Time = data.Time,
                    UseA = data.UseA,
                    UseV = data.UseV,
                });
            }
        }

        public static DataTable ConvertListToDataTable(List<Excel> Excel)
        {
            // New table.
            DataTable table = new DataTable();

            // Get max columns.
            int columns = 3;
            string[] name = new string[] { "เวลา", "การใช้แรงดัน (V)", "การใช้กระแส (A)" };
            // Add columns.
            for (int i = 0; i < columns; i++)
            {
                table.Columns.Add(name[i]);
            }

            // Add rows. and data in columns
            foreach (var array in Excel)
            {
                DataRow dr = table.NewRow();
                dr[0] = array.Time;
                dr[1] = array.UseV;
                dr[2] = array.UseA;
                table.Rows.Add(dr);
            }

            return table;
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void chart1_Click(object sender, EventArgs e)
        {

        }
    }
}
