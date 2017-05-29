using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Windows.Forms.DataVisualization.Charting;
using Excel = Microsoft.Office.Interop.Excel;

namespace ZubApproxProject
{
    public partial class Form1 : Form
    {
        Database dbconn;
        // double[] xArr;
        List<DateTime> xArr;
        List<double> yArr;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            servTB.Text = "46.101.139.249";
            userTB.Text = "zubUser";
            passTB.Text = "zubUser";
            bdTB.Text = "zubDB";
        }

        private void ConndbBtn_Click(object sender, EventArgs e)
        {
            dbconn = new Database(servTB.Text, userTB.Text, passTB.Text, bdTB.Text);
            List<string> tables = dbconn.ShowTables();
            comboBox1.Items.Clear();
            for (int i = 0; i < tables.Count; i++)
            {
                comboBox1.Items.Add(tables[i]);
            }
            comboBox1.SelectedIndexChanged += ComboBox1_SelectedIndexChanged; // подписываемся на событие выбора таблицы
            comboBox2.SelectedIndexChanged += ComboBoxBoth_SelectedIndexChanged; // подписываемся на событие выбора X значений
            comboBox3.SelectedIndexChanged += ComboBoxBoth_SelectedIndexChanged; // подписываемся на событие выбора У значений
    //        comboBox4.SelectedIndexChanged += ComboBox4_SelectedIndexChanged; // подписываемся на событие выбора для кого Херста вычисляе

            ConndbBtn.BackColor = Color.LightGreen;
            ConndbBtn.Text = "Подключено";
        }
     

        private void ComboBoxBoth_SelectedIndexChanged(object sender, EventArgs e) // заполняем X и Y
        {
            /* if (comboBox2.Text != "" && comboBox3.Text != "")
             {
                 xArr = new double[dgv1.Rows.Count - 1];
                 yArr = new double[dgv1.Rows.Count - 1];
                 for (int i = 0; i < dgv1.Rows.Count - 1; i++)
                 {
                     if (comboBox2.SelectedIndex >= 0)
                     {
                         xArr[i] = Convert.ToDouble(dgv1.Rows[i].Cells[comboBox2.SelectedIndex].Value);
                     }
                     if (comboBox3.SelectedIndex >= 0)
                     {
                         yArr[i] = Convert.ToDouble(dgv1.Rows[i].Cells[comboBox3.SelectedIndex].Value);
                     }
                 }

             } */
        }
        private void addColumn(string ColumnName)
        {
            var col1 = new DataGridViewTextBoxColumn();
            col1.HeaderText = ColumnName;
            col1.Name = ColumnName;
            //   dgvProg.Columns.Add(col1);
        }
        private void DgvProgClearColumn()
        {
            //   dgvProg.Columns.Clear();
        }

        private void ComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataTable dt = dbconn.SelectTable(comboBox1.Items[comboBox1.SelectedIndex].ToString());
            dgv1.DataSource = dt;

            comboBox2.Items.Clear();
            comboBox3.Items.Clear();
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                comboBox2.Items.Add(dt.Columns[i].ColumnName);
                comboBox3.Items.Add(dt.Columns[i].ColumnName);
            }
        }

        private void button1_Click(object sender, EventArgs e) // Построение линейного графика 
        {
            DrawLinearGraph(0);
        }
        private void DrawLinearGraph(int numSeries)
        {
            if (comboBox2.Text != "" || comboBox3.Text != "")
            {
                for (int i = 0; i < chart1.Series.Count; i++)
                {
                    if (i == 2 || i == 4 || i == 1)
                        chart1.Series[i].Points.Clear();
                }

                xArr = new List<DateTime>();
                yArr = new List<double>();
                //xArr = new DateTime[dgv1.Rows.Count - 1]; 
                //yArr = new double[dgv1.Rows.Count - 1];
                for (int i = 0; i < dgv1.Rows.Count - 1; i++)
                {
                    if (comboBox2.SelectedIndex >= 0)
                    {
                        try
                        {
                            if (dgv1.Rows[i].Cells[comboBox2.SelectedIndex].Value is DBNull)
                            {
                                Console.WriteLine("Хреновая");
                            }
                            else
                            {
                                //xArr[i] = Convert.ToDateTime(dgv1.Rows[i].Cells[comboBox2.SelectedIndex].Value);
                                xArr.Add(Convert.ToDateTime(dgv1.Rows[i].Cells[comboBox2.SelectedIndex].Value));
                            }
                        }
                        catch
                        {
                            MessageBox.Show("Хреновая строка мудень. Удали пустые строки.");
                        }
                    }
                    if (comboBox3.SelectedIndex >= 0)
                    {
                        try
                        {
                            if (dgv1.Rows[i].Cells[comboBox3.SelectedIndex].Value is DBNull)
                            {
                                Console.WriteLine("Хреновая");
                            }
                            else
                            {
                                // yArr[i] = Convert.ToDouble(dgv1.Rows[i].Cells[comboBox3.SelectedIndex].Value);
                                yArr.Add(Convert.ToDouble(dgv1.Rows[i].Cells[comboBox3.SelectedIndex].Value));
                            }
                        }
                        catch
                        {
                            MessageBox.Show("Хреновая строка мудень. Удали пустые строки.");
                        }
                    }
                }


                chart1.Series[numSeries].XValueType = ChartValueType.DateTime;
                chart1.ChartAreas[0].AxisX.LabelStyle.Format = "yyyy-MM-dd";
                chart1.ChartAreas[0].AxisX.Interval = 3;
                chart1.ChartAreas[0].AxisX.IntervalType = DateTimeIntervalType.Months;
                chart1.ChartAreas[0].AxisX.IntervalOffset = 1;

                chart1.Series[numSeries].XValueType = ChartValueType.DateTime;
                DateTime minDate = new DateTime(xArr[0].Year, xArr[0].Month, xArr[0].Day).AddSeconds(-1);
                DateTime maxDate = new DateTime(xArr[xArr.Count - 1].Year, xArr[xArr.Count - 1].Month, xArr[xArr.Count - 1].Day); // or DateTime.Now;
                chart1.ChartAreas[0].AxisX.Minimum = minDate.ToOADate();
                chart1.ChartAreas[0].AxisX.Maximum = maxDate.ToOADate();

                // chart1.ChartAreas[0].AxisY.Minimum = yArr.Min();
                double currentMin = yArr[0];
                for (int i = 0; i < yArr.Count; i++)
                {
                    if (yArr[i] != 0 && yArr[i] < currentMin)
                    {
                        currentMin = yArr[i];
                    }
                }
                chart1.ChartAreas[0].AxisY.Minimum = currentMin;
                chart1.ChartAreas[0].AxisY.Maximum = yArr.Max();

                ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            }
            else
            {
                MessageBox.Show("Выберите столбцы X и Y");
                return;
            }

            chart1.Series[numSeries].Points.Clear();

            chart1.Series[numSeries].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line; // тут сами поизменяет/повыбирайте тип вывода графика
            chart1.Series[numSeries].Color = Color.Blue;
            for (int i = 0; i < xArr.Count; i++)
            {
                if (yArr[i] != 0)
                {
                    chart1.Series[numSeries].Points.AddXY(xArr[i], yArr[i]);
                }
            }

            /*     var s = new Series();
                 s.ChartType = SeriesChartType.Line;

                 var d = new DateTime(2013, 04, 01);

                 s.Points.AddXY(d, 3);
                 s.Points.AddXY(d.AddMonths(-1), 2);
                 s.Points.AddXY(d.AddMonths(-2), 1);
                 s.Points.AddXY(d.AddMonths(-3), 4);

                 chart1.Series.Clear();
                 chart1.Series.Add(s);


                 chart1.Series[0].XValueType = ChartValueType.DateTime;
                 chart1.ChartAreas[0].AxisX.LabelStyle.Format = "yyyy-MM-dd";
                 chart1.ChartAreas[0].AxisX.Interval = 1;
                 chart1.ChartAreas[0].AxisX.IntervalType = DateTimeIntervalType.Months;
                 chart1.ChartAreas[0].AxisX.IntervalOffset = 1;

                 chart1.Series[0].XValueType = ChartValueType.DateTime;
                 DateTime minDate = new DateTime(2013, 01, 01).AddSeconds(-1);
                 DateTime maxDate = new DateTime(2013, 05, 01); // or DateTime.Now;
                 chart1.ChartAreas[0].AxisX.Minimum = minDate.ToOADate();
                 chart1.ChartAreas[0].AxisX.Maximum = maxDate.ToOADate(); */
        }

        private void AkimInterBtn_Click(object sender, EventArgs e) // Построение сплайна Акима 
        {
            /*   if (comboBox2.Text != "" || comboBox3.Text != "")
               {
                   xArr = new DateTime[dgv1.Rows.Count - 1];
                   yArr = new double[dgv1.Rows.Count - 1];
                   for (int i = 0; i < dgv1.Rows.Count - 1; i++)
                   {
                       if (comboBox2.SelectedIndex >= 0)
                       {
                           xArr[i] = Convert.ToDateTime(dgv1.Rows[i].Cells[comboBox2.SelectedIndex].Value);
                       }
                       if (comboBox3.SelectedIndex >= 0)
                       {
                           yArr[i] = Convert.ToDouble(dgv1.Rows[i].Cells[comboBox3.SelectedIndex].Value);
                       }
                   }
               }
               else
               {
                   MessageBox.Show("Выберите столбцы X и Y");
                   return;
               }

               chart1.Series[1].Points.Clear();

               var interAkima = MathNet.Numerics.Interpolation.CubicSpline.InterpolateAkima(xArr, yArr);

               chart1.Series[1].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line; // тут сами поизменяет/повыбирайте тип вывода графика
               chart1.Series[1].Color = Color.Orange;

               double step = 0.25;
               for (var x = xArr[0]; x <= xArr[xArr.Length - 1]; x += step)
               {
                   var interY = interAkima.Interpolate(x);
                   chart1.Series[1].Points.AddXY(x, interY);
               } */
            //MathNet.Numerics.LinearRegression.
          //  chart1.DataManipulator.FinancialFormula(System.Windows.Forms.DataVisualization.Charting.FinancialFormula.
            
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void VolatBtn_Click(object sender, EventArgs e)
        {            
           chart1.DataManipulator.FinancialFormula(System.Windows.Forms.DataVisualization.Charting.FinancialFormula.VolatilityChaikins, textBox1.Text, "Линейный:Y,Линейный2:Y", "Волатильность:Y");
            chart1.ChartAreas[0].AxisY.Minimum = double.NaN;
            chart1.ChartAreas[0].AxisY.Maximum = double.NaN;
            //херота
            for (int i = 0; i < chart1.Series.Count; i++)
            {
                if(i!=2)
                chart1.Series[i].Points.Clear();
            }
        }
        
        private void button2_Click(object sender, EventArgs e)
        {
            //   try
            // {
            OpenFileDialog openfile1 = new OpenFileDialog();
            if (openfile1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                nameSheets = GetSheetNameFromFile(openfile1);
                comboBox4.Items.Clear();
                for (int i = 0; i < nameSheets.Length; i++) // заполняем по какому значению можно вычислять Херста
                {
                    comboBox4.Items.Add(nameSheets[i]);
                }

            }
        }
        private string[] GetSheetNameFromFile(OpenFileDialog openfile1)
        {
            DataSource = openfile1.FileName;
            string pathconn = "Provider = Microsoft.ACE.OLEDB.12.0; Data source=" + openfile1.FileName + ";Extended Properties=\"Excel 8.0;HDR= yes;\";";
            OleDbConnection conn = new OleDbConnection(pathconn);
            //OleDbDataAdapter MyDataAdapter = new OleDbDataAdapter("Select * from [Лист1$]", conn);
            DataTable dt = new DataTable();
            //MyDataAdapter.Fill(dt);
            conn.Open();
            dt = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

            if (dt == null)
            {
                return null;
            }

            String[] excelSheets = new String[dt.Rows.Count];
            int i = 0;

            // Add the sheet name to the string array.
            foreach (DataRow row in dt.Rows)
            {
                excelSheets[i] = row["TABLE_NAME"].ToString();
                i++;
            }

            // Loop through all of the sheets if you want too...
            for (int j = 0; j < excelSheets.Length; j++)
            {
                // Query each excel sheet.
                excelSheets[j] = excelSheets[j].Trim("'".ToCharArray());
            }
            conn.Close();
            return excelSheets;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (comboBox2.Text != "" || comboBox3.Text != "")
            {
                for (int i = 0; i < chart1.Series.Count; i++)
                {
                    if (i == 2 || i == 4 || i == 1)
                        chart1.Series[i].Points.Clear();
                }

                xArr = new List<DateTime>();
                yArr = new List<double>();
                //xArr = new DateTime[dgv1.Rows.Count - 1]; 
                //yArr = new double[dgv1.Rows.Count - 1];
                for (int i = 0; i < dgv1.Rows.Count - 1; i++)
                {
                    if (comboBox2.SelectedIndex >= 0)
                    {
                        try
                        {
                            if (dgv1.Rows[i].Cells[comboBox2.SelectedIndex].Value is DBNull)
                            {
                                Console.WriteLine("Хреновая");
                            }
                            else
                            {
                                //xArr[i] = Convert.ToDateTime(dgv1.Rows[i].Cells[comboBox2.SelectedIndex].Value);
                                xArr.Add(Convert.ToDateTime(dgv1.Rows[i].Cells[comboBox2.SelectedIndex].Value));
                            }
                        }
                        catch
                        {
                            MessageBox.Show("Хреновая строка мудень. Удали пустые строки.");
                        }
                    }
                    if (comboBox3.SelectedIndex >= 0)
                    {
                        try
                        {
                            if (dgv1.Rows[i].Cells[comboBox3.SelectedIndex].Value is DBNull)
                            {
                                Console.WriteLine("Хреновая");
                            }
                            else
                            {
                                // yArr[i] = Convert.ToDouble(dgv1.Rows[i].Cells[comboBox3.SelectedIndex].Value);
                                yArr.Add(Convert.ToDouble(dgv1.Rows[i].Cells[comboBox3.SelectedIndex].Value));
                            }
                        }
                        catch
                        {
                            MessageBox.Show("Хреновая строка мудень. Удали пустые строки.");
                        }
                    }
                }


                chart1.Series[3].XValueType = ChartValueType.DateTime;
                chart1.ChartAreas[0].AxisX.LabelStyle.Format = "yyyy-MM-dd";
                chart1.ChartAreas[0].AxisX.Interval = 3;
                chart1.ChartAreas[0].AxisX.IntervalType = DateTimeIntervalType.Months;
                chart1.ChartAreas[0].AxisX.IntervalOffset = 1;

                chart1.Series[3].XValueType = ChartValueType.DateTime;
                DateTime minDate = new DateTime(xArr[0].Year, xArr[0].Month, xArr[0].Day).AddSeconds(-1);
                DateTime maxDate = new DateTime(xArr[xArr.Count - 1].Year, xArr[xArr.Count - 1].Month, xArr[xArr.Count - 1].Day); // or DateTime.Now;
                chart1.ChartAreas[0].AxisX.Minimum = minDate.ToOADate();
                chart1.ChartAreas[0].AxisX.Maximum = maxDate.ToOADate();

                // chart1.ChartAreas[0].AxisY.Minimum = yArr.Min();
                double currentMin = yArr[0];
                for (int i = 0; i < yArr.Count; i++)
                {
                    if (yArr[i] != 0 && yArr[i] < currentMin)
                    {
                        currentMin = yArr[i];
                    }
                }
                chart1.ChartAreas[0].AxisY.Minimum = currentMin;
                chart1.ChartAreas[0].AxisY.Maximum = yArr.Max();

                ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            }
            else
            {
                MessageBox.Show("Выберите столбцы X и Y");
                return;
            }

            chart1.Series[3].Points.Clear();

            chart1.Series[3].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line; // тут сами поизменяет/повыбирайте тип вывода графика
            chart1.Series[3].Color = Color.Blue;
            for (int i = 0; i < xArr.Count; i++)
            {
                if (yArr[i] != 0)
                {
                    chart1.Series[3].Points.AddXY(xArr[i], yArr[i]);
                }
            }
        }

        string DataSource = "";
        string[] nameSheets;
        private void comboBox4_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            string pathconn = "Provider = Microsoft.ACE.OLEDB.12.0; Data source=" + DataSource + ";Extended Properties=\"Excel 8.0;HDR= yes;\";";
            OleDbConnection conn = new OleDbConnection(pathconn);
            // conn.Open();
            OleDbDataAdapter MyDataAdapter = new OleDbDataAdapter($"Select * from [{nameSheets[comboBox4.SelectedIndex]}]", conn);
            DataTable dt = new DataTable();
            MyDataAdapter.Fill(dt);

            dgv1.DataSource = dt.DefaultView;

            comboBox2.Text = "";
            comboBox3.Text = "";
            comboBox2.Items.Clear();
            comboBox3.Items.Clear();
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                comboBox2.Items.Add(dt.Columns[i].ColumnName);
                comboBox3.Items.Add(dt.Columns[i].ColumnName);
            }
        }
        private void button4_Click(object sender, EventArgs e) // среднеквадратичное отклонение
        {
            if (chart1.Series[0].Points.Count > 0)
            {
                DrawSHV();
            }
            else
            {
                DrawLinearGraph(0);
                DrawSHV();
            }
            PopulateDGV2(4);
        }
        private void DrawSHV()
        {
            try
            {
                chart1.DataManipulator.FinancialFormula(FinancialFormula.StandardDeviation, textBox2.Text, "Линейный:Y", "Стндртклн:Y");
                chart1.ChartAreas[0].AxisY.Minimum = double.NaN;
                chart1.ChartAreas[0].AxisY.Maximum = double.NaN;
                //херота
                for (int i = 0; i < chart1.Series.Count; i++)
                {
                    if (i != 4)
                    {
                        chart1.Series[i].Points.Clear();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}");
            }
        }
        private void PopulateDGV2(int numSeries)
        {
            dgv2.Columns.Insert(0, new DataGridViewTextBoxColumn());
            dgv2.Columns.Insert(1, new DataGridViewTextBoxColumn());

            dgv2.Rows.Clear();
            for (int i=0; i < chart1.Series[numSeries].Points.Count;i++)
            {
                var pp = chart1.Series[numSeries].Points;
                dgv2.Rows.Add(xArr[i],chart1.Series[numSeries].Points[i].YValues[0].ToString());
            }
        }
        private void button5_Click(object sender, EventArgs e) // логарифмический масштаб
        {
            if (comboBox2.Text != "" || comboBox3.Text != "")
            {
                for (int i = 0; i < chart1.Series.Count; i++)
                {
                    if (i == 2 || i == 4 || i == 1)
                        chart1.Series[i].Points.Clear();
                }

                xArr = new List<DateTime>();
                yArr = new List<double>();
                //xArr = new DateTime[dgv1.Rows.Count - 1]; 
                //yArr = new double[dgv1.Rows.Count - 1];
                for (int i = 0; i < dgv1.Rows.Count - 1; i++)
                {
                    if (comboBox2.SelectedIndex >= 0)
                    {
                        try
                        {
                            if (dgv1.Rows[i].Cells[comboBox2.SelectedIndex].Value is DBNull)
                            {
                                Console.WriteLine("Хреновая");
                            }
                            else
                            {
                                //xArr[i] = Convert.ToDateTime(dgv1.Rows[i].Cells[comboBox2.SelectedIndex].Value);
                                xArr.Add(Convert.ToDateTime(dgv1.Rows[i].Cells[comboBox2.SelectedIndex].Value));
                            }
                        }
                        catch
                        {
                            MessageBox.Show("Хреновая строка мудень. Удали пустые строки.");
                        }
                    }
                    if (comboBox3.SelectedIndex >= 0)
                    {
                        try
                        {
                            if (dgv1.Rows[i].Cells[comboBox3.SelectedIndex].Value is DBNull)
                            {
                                Console.WriteLine("Хреновая");
                            }
                            else
                            {
                                // yArr[i] = Convert.ToDouble(dgv1.Rows[i].Cells[comboBox3.SelectedIndex].Value);
                                yArr.Add(Convert.ToDouble(dgv1.Rows[i].Cells[comboBox3.SelectedIndex].Value));
                            }
                        }
                        catch
                        {
                            MessageBox.Show("Хреновая строка мудень. Удали пустые строки.");
                        }
                    }
                }


                chart1.Series[0].XValueType = ChartValueType.DateTime;
                chart1.ChartAreas[0].AxisX.LabelStyle.Format = "yyyy-MM-dd";
                chart1.ChartAreas[0].AxisX.Interval = 3;
                chart1.ChartAreas[0].AxisX.IntervalType = DateTimeIntervalType.Months;
                chart1.ChartAreas[0].AxisX.IntervalOffset = 1;

                chart1.Series[0].XValueType = ChartValueType.DateTime;
                DateTime minDate = new DateTime(xArr[0].Year, xArr[0].Month, xArr[0].Day).AddSeconds(-1);
                DateTime maxDate = new DateTime(xArr[xArr.Count - 1].Year, xArr[xArr.Count - 1].Month, xArr[xArr.Count - 1].Day); // or DateTime.Now;
                chart1.ChartAreas[0].AxisX.Minimum = minDate.ToOADate();
                chart1.ChartAreas[0].AxisX.Maximum = maxDate.ToOADate();

                // chart1.ChartAreas[0].AxisY.Minimum = yArr.Min();
                double currentMin = yArr[0];
                for (int i = 0; i < yArr.Count; i++)
                {
                    if (yArr[i] != 0 && yArr[i] < currentMin)
                    {
                        currentMin = yArr[i];
                    }
                }
             //   chart1.ChartAreas[0].AxisY.Minimum = currentMin;
              //  chart1.ChartAreas[0].AxisY.Maximum = yArr.Max();

                ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            }
            else
            {
                MessageBox.Show("Выберите столбцы X и Y");
                return;
            }

            chart1.Series[0].Points.Clear();

            chart1.Series[0].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line; // тут сами поизменяет/повыбирайте тип вывода графика
            chart1.Series[0].Color = Color.Blue;
            for (int i = 0; i < xArr.Count; i++)
            {
                if (yArr[i] != 0)
                {
                    chart1.Series[0].Points.AddXY(xArr[i], Math.Log(yArr[i]));
                    Console.WriteLine(Math.Log(yArr[i]));
                }
            }
            chart1.ChartAreas[0].AxisX.Minimum = double.NaN;
            chart1.ChartAreas[0].AxisY.Maximum = double.NaN;
            chart1.ChartAreas[0].AxisY.Minimum = double.NaN;
            chart1.ChartAreas[0].AxisY.Maximum = double.NaN;

            PopulateDGV2(0);
            /* if (button5.Text == "Вкл")
             {
                 chart1.ChartAreas[0].AxisX.Minimum = double.NaN;
                 chart1.ChartAreas[0].AxisY.Maximum = double.NaN;
                 chart1.ChartAreas[0].AxisY.Minimum = double.NaN;
                 chart1.ChartAreas[0].AxisY.Maximum = double.NaN;

                 // chart1.ChartAreas[0].AxisX.IsLogarithmic = true;
                 chart1.ChartAreas[0].AxisY.IsLogarithmic = true;
                 button5.Text = "Выкл";
             }
             else
             {
                 chart1.ChartAreas[0].AxisX.Minimum = double.NaN;
                 chart1.ChartAreas[0].AxisY.Maximum = double.NaN;
                 chart1.ChartAreas[0].AxisY.Minimum = double.NaN;
                 chart1.ChartAreas[0].AxisY.Maximum = double.NaN;

                 // chart1.ChartAreas[0].AxisX.IsLogarithmic = true;
                 chart1.ChartAreas[0].AxisY.IsLogarithmic = false;
                 button5.Text = "Вкл";
             } */
        }

        private void button6_Click(object sender, EventArgs e) // garch model 
        {
            //int q = 0;

       
            // Extreme.Mathematics.Vector<double> vecY = new Extreme.Mathematics.Vector<double>();
            /*  double[] yTrueArr = new double[yArr.Count];
              for(int i=0; i<yArr.Count;i++)
              {
                  if (yArr[i] != null || yArr[i] != 0.0)
                  {
                      yTrueArr[i] = yArr[i];
                  }
              } */


            //Extreme.Mathematics.Vector<double> vecY = yTrueArr;



            /* GarchModel model1 = new GarchModel(0, 2);
             model1.Compute();

             GarchModel model2 = new GarchModel(vecY,1,1);

             model2.InnovationDistribution = GarchInnovationDistribution.StudentTWithFixedDegreesOfFreedom;
             model2.StudentTDegreesOfFreedom = 10;


             Console.WriteLine("Constant: {0}", model2.Parameters[0].Value);
             Console.WriteLine("ARCH(1): {0}", model2.Parameters[1].Value);
             Console.WriteLine("GARCH(1): {0}", model2.Parameters[2].Value); */
        }

        private void button7_Click(object sender, EventArgs e) // Сохранить в таблицу
        {
            if (nametblTxt.Text.Length == 0 || dbconn == null)
            {
                MessageBox.Show("Введите название таблицы. И подключитесь в бд.");
            }
            else
            {
                List<List<string>> columnData = new List<List<string>>();
                List<string> columnName = new List<string>();
                for (int i = 0; i < dgv1.ColumnCount; i++)
                {
                    columnName.Add(dgv1.Columns[i].Name);
                }
                for (int x = 0; x < columnName.Count; x++)
                {
                    List<string> addCol = new List<string>();
                    columnData.Add(addCol);
                    for (int i = 0; i < dgv1.Rows.Count - 1; i++)
                    {
                        columnData[x].Add(dgv1.Rows[i].Cells[x].Value.ToString());
                    }
                }
                bool ka = dbconn.CreateTableAndInsertData(nametblTxt.Text, columnName, columnData);
                if (ka == true)
                {
                    MessageBox.Show("Данные успешно загружены");
                }
                // chart1.Series[0].Points;
                //dbconn.CreateTableAndInsertData(nametblTxt.Text,)
            }

        }

        private void button6_Click_1(object sender, EventArgs e) // сохранить в файл
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "MS excel|*.xlsx|MS excel|*.xls";
            saveFileDialog1.Title = "Save an Table File";
            saveFileDialog1.ShowDialog();

            // If the file name is not an empty string open it for saving.
            if (saveFileDialog1.FileName != "")
            {


                Excel.Application excelapp = new Excel.Application();
                Excel.Workbook workbook = excelapp.Workbooks.Add();
                Excel.Worksheet worksheet = workbook.ActiveSheet;

                int w = 1;
                foreach (DataGridViewColumn column in dgv1.Columns)
                {
                    worksheet.Rows[1].Columns[w] = column.HeaderText;
                    w++;
                }
                /*   worksheet.Rows[1].Columns[1] = "ФИО";
                   worksheet.Rows[1].Columns[2] = "Название теста";
                   worksheet.Rows[1].Columns[3] = "Процент правильных ответов";
                   worksheet.Rows[1].Columns[4] = "Время выполнения теста";
                   worksheet.Rows[1].Columns[5] = "Дата прохождения теста";*/
                //int q = 1;
                for (int i = 1; i < dgv1.RowCount + 1; i++)
                {
                    for (int j = 1; j < dgv1.ColumnCount + 1; j++)
                    {
                        Console.WriteLine($"J: {j}");
                        worksheet.Rows[i + 1].Columns[j] = dgv1.Rows[i - 1].Cells[j - 1].Value;
                        // q++;
                    }
                    //  q = 1;

                }

                excelapp.AlertBeforeOverwriting = false;
                Console.WriteLine($"{Application.StartupPath}\\Итоговые результаты.xlsx");
                workbook.SaveAs($"{saveFileDialog1.FileName}");
                // workbook.SaveAs($"{Application.StartupPath}\\kjk.xlsx");
                excelapp.Quit();
            }
        }
    }
}
