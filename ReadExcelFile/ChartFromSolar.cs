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
    public partial class ChartFromSolar : Form
    {
        List<Excel> fromExcel = new List<Excel>();
        public ChartFromSolar()
        {
            InitializeComponent();
            chart1.Visible = false;
            dataGridView1.Visible = false;
            panel1.Visible = false;
        }

        private void ChartFromSolar_Load(object sender, EventArgs e)
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
                chart1.Visible = true;
                dataGridView1.Visible = false;
                panel1.Visible = true;
                int i = 0;
                textBox1.Clear();
                textBox2.Clear();
                dataGridView1.Visible = false;
                panel1.Visible = true;
                chart1.Visible = true;
                chart1.ChartAreas[0].AxisX.Title = "เวลา";
                chart1.ChartAreas[0].AxisY.Title = "แรงดัน / กระแส";

                Axis xaxis = chart1.ChartAreas[0].AxisX;
                xaxis.IntervalType = DateTimeIntervalType.Hours;
                xaxis.Interval = 1440;

                //Loop Remove point in chart
                foreach (var series in chart1.Series)
                {
                    series.Points.Clear();
                }

                foreach (var data in fromExcel)
                {
                    //Gen Graph Series Volt
                    chart1.Series["Volt"].Points.AddXY(data.Time, data.SolarCharV);
                    chart1.Series["Volt"].Points[i].MarkerStyle = MarkerStyle.Circle;
                    chart1.Series["Volt"].Points[i].MarkerSize = (data.SolarCharV.Equals(data.SolarCharA) ? 5 : 3);
                    chart1.Series["Volt"].Points[i].MarkerColor = Color.Blue;
                    //Gen Graph Series Am
                    chart1.Series["Ampere"].Points.AddXY(data.Time, data.SolarCharA);
                    chart1.Series["Ampere"].Points[i].MarkerStyle = MarkerStyle.Circle;
                    chart1.Series["Ampere"].Points[i].MarkerSize = 3;
                    chart1.Series["Ampere"].Points[i].MarkerColor = Color.Orange;
                    i++;
                    // UseV += Convert.ToDouble(data.UseV);
                    // UseA += Convert.ToDouble(data.UseA);
                }

                string argVolt = (fromExcel.Sum(a => Convert.ToDouble(a.SolarCharV)) / fromExcel.Count()).ToString("#.##");
                string argAm = (fromExcel.Sum(a => Convert.ToDouble(a.SolarCharA)) / fromExcel.Count()).ToString("#.##");
                textBox1.AppendText(argVolt == "" ? "0.00" : argVolt);
                textBox2.AppendText(argAm == "" ? "0.00" : argAm);
            }
            catch (Exception ex)
            {
                MessageBox.Show("เกิดข้อผิดพลาด : " + ex.Message, "เกิดข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
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

        public void _FirstLoadForm()
        {
            chart1.Visible = false;
            dataGridView1.Visible = true;
            panel1.Visible = true;
            textBox1.Clear();
            textBox2.Clear();
            DataTable table = ConvertListToDataTable(fromExcel);
            dataGridView1.DataSource = table;
            dataGridView1.Columns[1].Width = 140;
            dataGridView1.Columns[2].Width = 140;

            string argVolt = (fromExcel.Sum(a => Convert.ToDouble(a.SolarCharV)) / fromExcel.Count()).ToString("#.##");
            string argAm = (fromExcel.Sum(a => Convert.ToDouble(a.SolarCharA)) / fromExcel.Count()).ToString("#.##");
            textBox1.AppendText(argVolt == "" ? "0.00" : argVolt);
            textBox2.AppendText(argAm == "" ? "0.00" : argAm);
        }

        public void recieve(List<Excel> excel)
        {
            foreach (var data in excel)
            {
                fromExcel.Add(new Excel
                {
                    Date = data.Date,
                    Time = data.Time,
                    SolarCharV = data.SolarCharV,
                    SolarCharA = data.SolarCharA,
                });
            }
        }

        public static DataTable ConvertListToDataTable(List<Excel> Excel)
        {
            // New table.
            DataTable table = new DataTable();

            // Get max columns.
            int columns = 3;
            string[] name = new string[] { "เวลา", "ชาร์จจากโซล่าเซลล์ (V)", "ชาร์จจากโซล่าเซลล์ (A)" };
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
                dr[1] = array.SolarCharV;
                dr[2] = array.SolarCharA;
                table.Rows.Add(dr);
            }
            return table;
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}
