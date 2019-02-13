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
    public partial class ShowAll : Form
    {
        List<Excel> fromExcel = new List<Excel>();
        public ShowAll()
        {
            InitializeComponent();
        }

        private void ShowAll_Load(object sender, EventArgs e)
        {
            MainMenu f1 = new MainMenu();
            try
            {
                
                int Indexchart = 0;
                //int Indexchart2 = 0;
                //int Indexchart3 = 0;
                //int Indexchart4 = 0;

                //use elec per day
                textBox1.Clear(); 
                textBox2.Clear();
                textBox1.AppendText(fromExcel.Average(x => Convert.ToDouble(x.UseV)).ToString("#.##") == "" ? "0.00" : fromExcel.Average(x => Convert.ToDouble(x.UseV)).ToString("#.##"));
                textBox2.AppendText(fromExcel.Average(x => Convert.ToDouble(x.UseA)).ToString("#.##") == "" ? "0.00" : fromExcel.Average(x => Convert.ToDouble(x.UseA)).ToString("#.##"));
                //House Char per day
                textBox6.Clear();
                textBox5.Clear();
                textBox6.AppendText(fromExcel.Average(x => Convert.ToDouble(x.HomeCharV)).ToString("#.##") == "" ? "0.00" : fromExcel.Average(x => Convert.ToDouble(x.HomeCharV)).ToString("#.##"));
                textBox5.AppendText(fromExcel.Average(x => Convert.ToDouble(x.HomeCharA)).ToString("#.##") == "" ? "0.00" : fromExcel.Average(x => Convert.ToDouble(x.HomeCharA)).ToString("#.##"));
                //Solar char per day
                textBox4.Clear();
                textBox3.Clear();
                textBox4.AppendText(fromExcel.Average(x => Convert.ToDouble(x.SolarCharV)).ToString("#.##") == "" ? "0.00" : fromExcel.Average(x => Convert.ToDouble(x.SolarCharV)).ToString("#.##"));
                textBox3.AppendText(fromExcel.Average(x => Convert.ToDouble(x.SolarCharA)).ToString("#.##") == "" ? "0.00" : fromExcel.Average(x => Convert.ToDouble(x.SolarCharA)).ToString("#.##"));
                //Distanc
                textBox7.Clear();
                textBox7.AppendText(fromExcel.Average(x => Convert.ToDouble(x.Distanc)).ToString("#.##") == "" ? "0.00" : fromExcel.Average(x => Convert.ToDouble(x.Distanc)).ToString("#.##"));

                //Chart1 use electric per Day
                chart1.ChartAreas[0].AxisX.Title = "เวลา";
                chart1.ChartAreas[0].AxisY.Title = "แรงดัน / กระแส";
                //Chart2 Houes Char Per Day
                chart2.ChartAreas[0].AxisX.Title = "เวลา";
                chart2.ChartAreas[0].AxisY.Title = "แรงดัน / กระแส";
                //Chart3 Solar char per Day
                chart3.ChartAreas[0].AxisX.Title = "เวลา";
                chart3.ChartAreas[0].AxisY.Title = "แรงดัน / กระแส";
                //Chart3 Distanc use per Day
                chart2.ChartAreas[0].AxisX.Title = "เวลา";
                chart2.ChartAreas[0].AxisY.Title = "กิโลเมตร";

                //Gen Chart1 use Electric per day
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
                    chart1.Series["Volt"].Points.AddXY(data.Time, data.UseV);
                    chart1.Series["Volt"].Points[Indexchart].MarkerStyle = MarkerStyle.Circle;
                    chart1.Series["Volt"].Points[Indexchart].MarkerSize = (data.UseV.Equals(data.UseA) ? 5 : 3);
                    chart1.Series["Volt"].Points[Indexchart].MarkerColor = Color.Blue;
                    //Gen Graph Series Am
                    chart1.Series["Ampere"].Points.AddXY(data.Time, data.UseA);
                    chart1.Series["Ampere"].Points[Indexchart].MarkerStyle = MarkerStyle.Circle;
                    chart1.Series["Ampere"].Points[Indexchart].MarkerSize = 3;
                    chart1.Series["Ampere"].Points[Indexchart].MarkerColor = Color.Orange;
                    Indexchart++;
                }
                Indexchart = 0;
                //Gen Chart2 House char per day
                Axis xaxis2 = chart2.ChartAreas[0].AxisX;
                xaxis2.IntervalType = DateTimeIntervalType.Hours;
                xaxis2.Interval = 1440;
                //Loop Remove point in chart
                foreach (var series in chart2.Series)
                {
                    series.Points.Clear();
                }

                //Loop Add point in chart
                foreach (var data in fromExcel)
                {
                    //Gen Graph Series Volt
                    chart2.Series["Volt"].Points.AddXY(data.Time, data.HomeCharV);
                    chart2.Series["Volt"].Points[Indexchart].MarkerStyle = MarkerStyle.Circle;
                    chart2.Series["Volt"].Points[Indexchart].MarkerSize = (data.HomeCharV.Equals(data.HomeCharA) ? 5 : 3);
                    chart2.Series["Volt"].Points[Indexchart].MarkerColor = Color.Navy;
                    //Gen Graph Series Am
                    chart2.Series["Ampere"].Points.AddXY(data.Time, data.HomeCharA);
                    chart2.Series["Ampere"].Points[Indexchart].MarkerStyle = MarkerStyle.Circle;
                    chart2.Series["Ampere"].Points[Indexchart].MarkerSize = 3;
                    chart2.Series["Ampere"].Points[Indexchart].MarkerColor = Color.SandyBrown;
                    Indexchart++;
                }

                Indexchart = 0;
                //Gen Chart3 Solar char per day
                Axis xaxis3 = chart3.ChartAreas[0].AxisX;
                xaxis3.IntervalType = DateTimeIntervalType.Hours;
                xaxis3.Interval = 1440;
                //Loop Remove point in chart
                foreach (var series in chart3.Series)
                {
                    series.Points.Clear();
                }

                //Loop Add point in chart
                foreach (var data in fromExcel)
                {
                    //Gen Graph Series Volt
                    chart3.Series["Volt"].Points.AddXY(data.Time, data.SolarCharV);
                    chart3.Series["Volt"].Points[Indexchart].MarkerStyle = MarkerStyle.Circle;
                    chart3.Series["Volt"].Points[Indexchart].MarkerSize = (data.SolarCharV.Equals(data.SolarCharA) ? 5 : 3);
                    chart3.Series["Volt"].Points[Indexchart].MarkerColor = Color.Indigo;
                    //Gen Graph Series Am
                    chart3.Series["Ampere"].Points.AddXY(data.Time, data.SolarCharA);
                    chart3.Series["Ampere"].Points[Indexchart].MarkerStyle = MarkerStyle.Circle;
                    chart3.Series["Ampere"].Points[Indexchart].MarkerSize = 3;
                    chart3.Series["Ampere"].Points[Indexchart].MarkerColor = Color.Peru;
                    Indexchart++;
                }

                Indexchart = 0;
                //Gen Chart4 Distanc use per day
                Axis xaxis4 = chart4.ChartAreas[0].AxisX;
                xaxis4.IntervalType = DateTimeIntervalType.Hours;
                xaxis4.Interval = 1440;
                //Loop Remove point in chart
                foreach (var series in chart4.Series)
                {
                    series.Points.Clear();
                }

                //Loop Add point in chart
                foreach (var data in fromExcel)
                {
                    //Gen Graph Series KM
                    chart4.Series["KM"].Points.AddXY(data.Time, data.Distanc);
                    chart4.Series["KM"].Points[Indexchart].MarkerStyle = MarkerStyle.Circle;
                    chart4.Series["KM"].Points[Indexchart].MarkerSize = 3;
                    chart4.Series["KM"].Points[Indexchart].MarkerColor = Color.DarkGreen;                   
                    Indexchart++;
                }

                Indexchart = 0;

            }
            catch (Exception ex)
            {
                
                MessageBox.Show("เกิดข้อผิดพลาด : " + ex.Message, "เกิดข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Warning);

            }
        }

        public void recieve(List<Excel> excel)
        {
            foreach (var data in excel)
            {
                fromExcel.Add(new Excel
                {
                    Date = data.Date,
                    Time = data.Time,
                    HomeCharV = data.HomeCharV,
                    HomeCharA = data.HomeCharA,
                    SolarCharV = data.SolarCharV,
                    SolarCharA = data.SolarCharA,
                    UseV = data.UseV,
                    UseA = data.UseA,
                    Distanc = data.Distanc,
                });
            }
        }
    }
}
