using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;


namespace ReadExcelFile
{
    public partial class MainMenu : Form
    {
        List<Excel> fromExcel = new List<Excel>();
        public MainMenu()
        {
            InitializeComponent();
            button2.Enabled = false;
            button3.Enabled = false;
            button4.Enabled = false;
            button5.Enabled = false;
            button6.Enabled = false;
            button7.Enabled = false;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {

                string fname = "";
                OpenFileDialog fdlg = new OpenFileDialog();
                fdlg.Title = "Excel File Dialog";
                //fdlg.InitialDirectory = @"c:\nack\";
                fdlg.InitialDirectory = @"c:\";

                fdlg.Filter = "All files (*.*)|*.*|All files (*.*)|*.*";
                fdlg.FilterIndex = 2;
                fdlg.RestoreDirectory = true;
                if (fdlg.ShowDialog() == DialogResult.OK)
                {
                    fname = fdlg.FileName;
                }


                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(fname);
                Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;

                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;

                // dt.Column = colCount;                              
                dataGridView1.ColumnCount = colCount;
                dataGridView1.RowCount = rowCount;
                //Set Head DataGrid
                dataGridView1.Columns[0].HeaderText = "Day/Mount/Year";
                dataGridView1.Columns[1].HeaderText = "Hour/Min";
                dataGridView1.Columns[2].HeaderText = "House Char (A)";
                dataGridView1.Columns[3].HeaderText = "House Char (V)";
                dataGridView1.Columns[4].HeaderText = "Solar Char (A)";
                dataGridView1.Columns[5].HeaderText = "Solar Char (V)";
                dataGridView1.Columns[6].HeaderText = "Use (A)";
                dataGridView1.Columns[7].HeaderText = "Use (V)";
                dataGridView1.Columns[8].HeaderText = "Distanc (KM)";
                ///int countNumberRows = 0;

                fromExcel.Clear();
                for (int i = 1; i <= rowCount; i++)
                {
                    double time = double.Parse(xlRange.Cells[i, 2].Value2.ToString());
                    DateTime parseTime = DateTime.FromOADate(time);
                    for (int j = 1; j <= colCount; j++)
                    {

                        var s = xlRange.Cells[i, j].Value2.ToString();
                        //write the value to the Grid                      
                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                        {
                            dataGridView1.Rows[i - 1].Cells[j - 1].Value = !String.IsNullOrEmpty(xlRange.Cells[i, j].Value2.ToString()) ?
                                (j == 1 ? string.Format("{0:dd-MMM-yyyy}", xlRange.Cells[i, j].Value2.ToString()) :
                                j == 2 ? dataGridView1.Rows[i - 1].Cells[j - 1].Value = string.Format("{0: h:mm:ss tt}", parseTime) :
                                dataGridView1.Rows[i - 1].Cells[j - 1].Value = string.Format("{0:00}", xlRange.Cells[i, j].Value2.ToString()))
                                : string.Format("{0:00}", "0");
                        }
                        Console.Write(xlRange.Cells[i, j].Value2.ToString() + "\t");  
                        //// dataGridView1.Rows[(countNumberRows + 1)].HeaderCell.Value = (countNumberRows + 1).ToString();
                        ////add useful things here!                       
                    }

                    if (xlRange.Cells[i, 1].Value2.ToString() != "day/month/year")
                    {
                        fromExcel.Add(new Excel
                        {
                            Date = string.Format("{0:dd-MMM-yyyy}", xlRange.Cells[i, 1].Value2.ToString()),
                            Time = string.Format("{0: h:mm:ss tt}", parseTime),
                            HomeCharA = xlRange.Cells[i, 3].Value2.ToString(),
                            HomeCharV = xlRange.Cells[i, 4].Value2.ToString(),
                            SolarCharA = xlRange.Cells[i, 5].Value2.ToString(),
                            SolarCharV = xlRange.Cells[i, 6].Value2.ToString(),
                            UseA = xlRange.Cells[i, 7].Value2.ToString(),
                            UseV = xlRange.Cells[i, 8].Value2.ToString(),
                            Distanc = xlRange.Cells[i, 9].Value2.ToString(),
                        });
                    }

                    if (i == rowCount)
                    {
                        button2.Enabled = true;
                        button3.Enabled = true;
                        button4.Enabled = true;
                        button5.Enabled = true;
                        button6.Enabled = true;
                        button7.Enabled = true;
                        MessageBox.Show("อัพโหลดเสร็จสิ้น", "การอัพโหลด", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }

                //cleanup  
                GC.Collect();
                GC.WaitForPendingFinalizers();

                //rule of thumb for releasing com objects:  
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
            catch (Exception ex)
            {
                MessageBox.Show("กรุณาเลือกไฟล์ Excel !", "เกิดข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                MainMenu_Load(sender, e);
            }
        }

        //Btn use currency from battery
        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                UseElecFromBattery f2 = new UseElecFromBattery();
                if (fromExcel != null && fromExcel.Count() != 0)
                {
                    f2.recieve(fromExcel);
                    f2.ShowDialog(); // Shows Form2
                }
                else
                {
                    button1.PerformClick();
                    //button1_Click(new object(), new EventArgs());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("เกิดข้อผิดพลาด : " + ex.Message, "เกิดข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        //Btn chart electri from solar 
        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                ChartFromSolar f3 = new ChartFromSolar();
                //If List null and count != 0 call aother form
                if (fromExcel != null && fromExcel.Count() != 0)
                {
                    f3.recieve(fromExcel);
                    f3.ShowDialog(); // Shows Form3
                }
                else
                {
                    button1.PerformClick(); //click button1 for search file
                    //button1_Click(new object(), new EventArgs());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("เกิดข้อผิดพลาด : " + ex.Message, "เกิดข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        //Btn chart electri from Home
        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                ChartFromHome f4 = new ChartFromHome();
                if (fromExcel != null && fromExcel.Count() != 0)
                {
                    f4.recieve(fromExcel);
                    f4.ShowDialog();
                }
                else
                {
                    button1.PerformClick();
                    //button1_Click(new object(), new EventArgs());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("เกิดข้อผิดพลาด : " + ex.Message, "เกิดข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        //btn Show Distanc
        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                Distance f5 = new Distance();
                if (fromExcel != null && fromExcel.Count() != 0)
                {
                    f5.recieve(fromExcel);
                    f5.ShowDialog();
                }
                else
                {
                    //button1.PerformClick();
                    button1_Click(new object(), new EventArgs());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("เกิดข้อผิดพลาด : " + ex.Message, "เกิดข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        // btn Show all detail and grap
        private void button6_Click(object sender, EventArgs e)
        {
            try
            {
                ShowAll f6 = new ShowAll();
                if (fromExcel != null && fromExcel.Count() != 0)
                {
                    f6.recieve(fromExcel);
                    f6.ShowDialog();
                }
                else
                {
                    button1.PerformClick();
                    //button1_Click(new object(), new EventArgs());
                }
            }
            catch (Exception ex)
            {
                button1_Click(new object(), new EventArgs());
                MessageBox.Show("เกิดข้อผิดพลาด : " + ex.Message +" ; กรุณาเลือกค้นหาไฟล์", "เกิดข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            try
            {
                PowerAndEnergy f7 = new PowerAndEnergy();
                if (fromExcel != null && fromExcel.Count() != 0)
                {
                    f7.recieve(fromExcel);
                    f7.ShowDialog();
                }
                else
                {
                    button1.PerformClick();                    
                }
            }
            catch (Exception ex)
            {
                button1_Click(new object(), new EventArgs());
                MessageBox.Show("เกิดข้อผิดพลาด : " + ex.Message + " ; กรุณาเลือกค้นหาไฟล์", "เกิดข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Warning);

            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
        }

        private void MainMenu_Load(object sender, EventArgs e)
        {

        }        
    }
}
