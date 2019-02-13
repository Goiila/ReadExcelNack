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
    public partial class Distance : Form
    {
        List<Excel> fromExcel = new List<Excel>();
        public Distance()
        {
            InitializeComponent();
            chart1.Visible = false;
            dataGridView1.Visible = false;
            panel1.Visible = false;
        }

        private void Distance_Load(object sender, EventArgs e)
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
                dataGridView1.Visible = false;
                panel1.Visible = true;
                chart1.Visible = true;
                chart1.ChartAreas[0].AxisX.Title = "เวลา";
                chart1.ChartAreas[0].AxisY.Title = "กิโลเมตร";
                Axis xaxis = chart1.ChartAreas[0].AxisX;
                xaxis.IntervalType = DateTimeIntervalType.Hours;
                xaxis.Interval = 1440;
                //Loop Remove point in chart
                foreach (var series in chart1.Series)
                {
                    series.Points.Clear();
                }

                //Loop Add point in chart
                foreach (var data in fromExcel)
                {
                    //Gen Graph Series Volt
                    chart1.Series["KM"].Points.AddXY(data.Time, data.Distanc);
                    chart1.Series["KM"].Points[i].MarkerStyle = MarkerStyle.Circle;
                    chart1.Series["KM"].Points[i].MarkerSize = 3;
                    chart1.Series["KM"].Points[i].MarkerColor = Color.Blue;
                    //Gen Graph Series Am
                    i++;
                    // UseV += Convert.ToDouble(data.UseV);
                    // UseA += Convert.ToDouble(data.UseA);
                }

                string ArgDistance = (fromExcel.Sum(a => Convert.ToDouble(a.Distanc)) / fromExcel.Count()).ToString("#.##");
                textBox1.AppendText(ArgDistance == "" ? "0.00" : ArgDistance);
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
            DataTable table = ConvertListToDataTable(fromExcel);
            dataGridView1.DataSource = table;
            dataGridView1.Columns[1].Width = 160;
            string ArgDistance = (fromExcel.Sum(a => Convert.ToDouble(a.Distanc)) / fromExcel.Count()).ToString("#.##");
            textBox1.AppendText(ArgDistance == "" ? "0.00" : ArgDistance);
        }
        public void recieve(List<Excel> excel)
        {
            foreach (var data in excel)
            {
                fromExcel.Add(new Excel
                {
                    Date = data.Date,
                    Time = data.Time,
                    Distanc = data.Distanc,
                });
            }
        }

        public static DataTable ConvertListToDataTable(List<Excel> Excel)
        {
            // New table.
            DataTable table = new DataTable();

            // Get max columns.
            int columns = 2;
            string[] name = new string[] { "เวลา", "ระยะทาง (KM)" };
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
                dr[1] = array.Distanc;
                table.Rows.Add(dr);
            }
            return table;
        }

    }
}
