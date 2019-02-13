using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ReadExcelFile
{
    public partial class PowerAndEnergy : Form
    {  
        List<Excel> fromExcel = new List<Excel>();
     
        public PowerAndEnergy()
        {
            InitializeComponent();
        }

        private void PowerAndEnergy_Load(object sender, EventArgs e)
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
            List<CalculatePower> calPower = new List<CalculatePower>();
            List<ListCalPower> listCalPowers = new List<ListCalPower>();
            List<Excel> _2Excel = new List<Excel>();
            for(int i = 0; i < fromExcel.Count(); i++)
            {
                if (!fromExcel[i].UseA.Equals("0"))
                {                   
                    calPower.Add(new CalculatePower
                    {                       
                        Time = fromExcel[i].Time,
                        Power = Math.Pow(Convert.ToDouble(fromExcel[i].UseA), 2) * Convert.ToDouble(fromExcel[i].UseV),
                    });
                }
                
                if(calPower.Count() > 0 && fromExcel[i].UseA.Equals("0"))
                {
                    listCalPowers.Add(new ListCalPower {
                        calObj = calPower
                    });
                    calPower = new List<CalculatePower>();
                }
            }
            int number = 1;
            foreach(var row in listCalPowers)
            {
                double? power = row.calObj.Average(x => x.Power);
                double? energy = power * (double)(Convert.ToDateTime(row.calObj.LastOrDefault().Time) - Convert.ToDateTime(row.calObj.FirstOrDefault().Time)).TotalSeconds;
                _2Excel.Add(new Excel
                {
                    Time = number.ToString(),
                    Power = power.ToString().Length >= 4 && power.ToString().Length <= 6 ? String.Format("{0:0.00}", power / 1000) + " (KW)" : 
                    (power.ToString().Length >= 7 && power.ToString().Length <= 10 ? String.Format("{0:0.00}", power / 1000000) : String.Format("{0:0.00}", power) + " (W)"),
                    Energy = energy.ToString().Length >= 4 && energy.ToString().Length <= 6 ? String.Format("{0:0.00}", energy / 1000)+" (KJ)" : 
                    (energy.ToString().Length >= 7 && energy.ToString().Length <= 10 ? String.Format("{0:0.00}", energy / 1000000)+" (MJ)" : String.Format("{0:0.00}", energy) + " (J)"),
                });
                number++;
            }

            dataGridView1.Visible = true;
            //chart1.Visible = false;
            //panel1.Visible = true;
            //textBox1.Clear();
            //textBox2.Clear();
            ////Initialize dataGridView1
            //dataGridView1.Columns[0].HeaderText = "ครั้งที่";
            //dataGridView1.Columns[1].HeaderText = "กำลังไฟที่ใช้";
            //dataGridView1.Columns[2].HeaderText = "พลังงานที่ใช้";
            // Convert to DataTable.
            DataTable table = ConvertListToDataTable(_2Excel);
            dataGridView1.DataSource = table;
            dataGridView1.Columns[1].Width = 120;
            dataGridView1.Columns[2].Width = 120;

            //string ArgVolt = (fromExcel.Sum(a => Convert.ToDouble(a.UseV)) / fromExcel.Count()).ToString("#.##");
            //string ArgAm = (fromExcel.Sum(a => Convert.ToDouble(a.UseA)) / fromExcel.Count()).ToString("#.##");
            //textBox1.AppendText(ArgVolt == "" ? "0.00" : ArgVolt);
            //textBox2.AppendText(ArgAm == "" ? "0.00" : ArgAm);
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

        public static DataTable ConvertListToDataTable(List<Excel> Excel)
        {
            // New table.
            DataTable table = new DataTable();

            // Get max columns.
            int columns = 3;
            string[] name = new string[] { "ครั้งที่", "กำลังไฟที่ใช้", "พลังงานที่ใช้" };
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
                dr[1] = array.Power;
                dr[2] = array.Energy;
                table.Rows.Add(dr);
            }

            return table;
        }
        
    }
}
