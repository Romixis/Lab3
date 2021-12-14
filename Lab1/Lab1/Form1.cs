using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using org.mariuszgromada.math.mxparser;
using Excel = Microsoft.Office.Interop.Excel;

namespace Lab1
{
    public partial class Метод : Form
    {
        public static List<Point> XY = new List<Point>(); // Список для точек
        public string func;
        public string func2;
        string[,] list;
        public class Point // Сохраниние точек
        {
            public double x, y;
            public Point(double X, double Y)
            {
                this.x = X;
                this.y = Y;
            }
        }
        public Метод()
        {
            InitializeComponent();
            StartPosition = FormStartPosition.CenterScreen;  // Вывод формы по центру экрана
            label1.Text = "";
            label2.Text = "";
        }

        private void очиститьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Cl();
        }

        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void считатьСExcelToolStripMenuItem_Click(object sender, EventArgs e) // Считывание с Excel
        {
            Cl();

            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = "*.xls;*.xlsx";
            ofd.Filter = "файл Excel (Spisok.xlsx)|*.xlsx";
            ofd.Title = "Выберите файл базы данных";

            if (!(ofd.ShowDialog() == DialogResult.OK))
                MessageBox.Show("", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);

            Excel.Application ObjWorkExcel = new Excel.Application();
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(ofd.FileName);
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];

            var lastCell = ObjWorkSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell);
            int lastColumn = (int)lastCell.Column;
            int lastRow = (int)lastCell.Row;

            list = new string[lastRow, lastColumn];

            for (int j = 0; j < 2; ++j)
                for (int i = 0; i < lastRow; ++i)
                    list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text.ToString();

            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
            GC.Collect();
            for (int i = 0; i < list.GetLength(0); ++i)
            {
                Point xy = new Point(Convert.ToDouble(list[i, 0]), Convert.ToDouble(list[i, 1]));
                XY.Add(xy);
            }

            foreach (var p in XY)
            {
                dataGridView1.Rows.Add(p.x, p.y);
                chart1.Series[1].Points.AddXY(p.x, p.y);
            }
        }

        private void считатьСGoogleSheetsToolStripMenuItem_Click(object sender, EventArgs e) // Считывание с Google Sheets
        {
            try
            {
                Cl();

                string path = @"C:\Users\roma_\Desktop\progekt\Lab3\Lab1\2.xlsx";
                System.IO.File.Delete(path);

                string link = textBox1.Text;
                string qq = link.Replace("edit?usp=sharing", "export?format=xlsx");

                using (var client = new WebClient()) // Скачивание файла
                {
                    client.DownloadFile(new Uri(qq), path);
                }

                Excel.Application ObjExcel = new Excel.Application();

                Workbook ObjWorkBook = ObjExcel.Workbooks.Open(path); // Открываем книгу
                Worksheet ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1]; // Выбираем лист

                Range xRange = ObjWorkSheet.UsedRange.Columns[1]; // Первый столбец
                Range yRange = ObjWorkSheet.UsedRange.Columns[2]; // Второй столбец
                Array xCells = (Array)xRange.Cells.Value2;
                Array yCells = (Array)yRange.Cells.Value2;

                var lastCell = ObjWorkSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell);
                int lastColumn = (int)lastCell.Column;
                int lastRow = (int)lastCell.Row;

                list = new string[lastRow, lastColumn];

                for (int j = 0; j < 2; ++j)
                    for (int i = 0; i < lastRow; ++i)
                        list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text.ToString();

                ObjWorkBook.Close(false, Type.Missing, Type.Missing);
                ObjExcel.Quit();
                GC.Collect();
                for (int i = 0; i < list.GetLength(0); ++i)
                {
                    Point xy = new Point(Convert.ToDouble(list[i, 0]), Convert.ToDouble(list[i, 1]));
                    XY.Add(xy);
                }
                foreach (var p in XY)
                {
                    dataGridView1.Rows.Add(p.x, p.y);
                    chart1.Series[1].Points.AddXY(p.x, p.y);
                }
            }
            catch
            {
                MessageBox.Show("Ввидете ссылку с правами на чтение!", "Ошибка");
            }
        }

        public void MNK()
        {
            //переменные
            double sumX = 0; double sumY = 0; double sumXY = 0; double X2 = 0; double Y2 = 0; double X3 = 0; double X4 = 0; double x2y = 0;
            double count = XY.Count;

            double a, b, del1, delt2, delt3, delt4;

            foreach (var p in XY)
            {
                sumX += p.x;
                sumY += p.y;
                sumXY += p.x * p.y;
                X2 += p.x * p.x;
                Y2 += p.y * p.y;
                X3 += p.x * p.x * p.x;
                X4 += p.x * p.x * p.x * p.x;
                x2y += (p.x * p.x) * p.y;
            }

            a = Math.Round(LinearA(sumX, sumY, sumXY, X2, Y2, count), 4); // A и B для линейной
            b = Math.Round(LinearB(sumX, sumY, sumXY, X2, Y2, count), 4);

            del1 = Det1(X2, sumX, count, X3, X4);
            delt2 = Math.Round(Det2(X2, sumX, count, X3, X4, sumY, sumXY, x2y) / del1, 4);

            delt3 = Math.Round(((X2 * sumXY * X2) + (sumY * sumX * X4) + (count * X3 * x2y) - 
                (count * sumXY * X4) - (X2 * sumX * x2y) - (sumY * X3 * X2)) / del1, 4);

            delt4 = Math.Round(((X2 * X2 * x2y) + (sumX * sumXY * X4) + (sumY * X3 * X3) - 
                (sumY * X2 * X4) - (X2 * sumXY * X3) - (sumX * X3 * x2y)) / del1, 4);

            func = $"f(x) = {a} * x + {b}";
            func = func.Replace(",", ".");

            func2 = $"f(x) = {delt2} * x^2 + ({delt3}) * x + {delt4}";
            func2 = func2.Replace(",", ".");

            label1.Text = func;
            label2.Text = func2;

            graph();
        }
        public void graph()//отрисовка графа
        {
            double min = Double.MaxValue;
            double max = Double.MinValue;
            double step = 1;

            for (int i = 0; i < XY.Count; ++i)//поиск минимума и максимума
            {
                if (XY[i].x < min)
                    min = XY[i].x;
                if (XY[i].x > max)
                    max = XY[i].x;
            }

            int count = (int)Math.Ceiling((max - min) / step) + 1;

            double[] x = new double[count];
            double[] y = new double[count];
            double[] y1 = new double[count];

            for (int i = 0; i < count; ++i)
            {
                x[i] = min + step * i;
                y[i] = Math.Round(Lin(x[i]), 5);
                y1[i] = Math.Round(Gip(x[i]), 5);
            }

            chart1.Series[0].Points.DataBindXY(x, y);//отрисовка
            chart1.Series[2].Points.DataBindXY(x, y1);
        }

        private double Lin(double x)
        {
            double result = 0;
            Function f = new Function(func);
            string sklt = "f()";
            string fx = sklt.Insert(2, x.ToString());
            fx = fx.Replace(",", ".");
            Expression fxx = new Expression(fx, f);
            result = fxx.calculate();
            return result;
        }

        private double Gip(double x)
        {
            double result = 0;
            Function f = new Function(func2);
            string sklt = "f()";
            string fx = sklt.Insert(2, x.ToString());
            fx = fx.Replace(",", ".");
            Expression fxx = new Expression(fx, f);
            result = fxx.calculate();
            return result;
        }

        public double LinearA(double SumX, double SumY, double SumXY, double X2, double Y2, double count)
        {
            double a;
            a = (SumX * SumY - count * SumXY) / ((SumX * SumX) - count * X2);
            return a;
        }

        public double LinearB(double SumX, double SumY, double SumXY, double X2, double Y2, double count)
        {
            double b;
            b = (SumX * SumXY - X2 * SumY) / ((SumX * SumX) - count * X2);
            return b;
        }

        public double Det1(double X2, double SumX, double count, double X3, double X4)
        {
            double[,] qq = new double[3, 3] { { X2, SumX, count }, { X3, X2, SumX }, { X4, X3, X2 } };
            double det = qq[0, 0] * qq[1, 1] * qq[2, 2] + qq[0, 1] * qq[1, 2] * qq[2, 0] +
                qq[0, 2] * qq[1, 0] * qq[2, 1] - qq[0, 2] * qq[2, 2] * qq[2, 0] -
                qq[0, 0] * qq[1, 2] * qq[2, 1] - qq[0, 1] * qq[1, 0] * qq[2, 2];
            return det;
        }
        public double Det2(double X2, double SumX, double count, double X3, double X4, double sumY, double sumXY, double x2y) //замудренный расчет дельты
        {
            double[,] qq = new double[3, 3] { { sumY, SumX, count }, { sumXY, X2, SumX }, { x2y, X3, X2 } };
            double det = qq[0, 0] * qq[1, 1] * qq[2, 2] + qq[0, 1] * qq[1, 2] * qq[2, 0] +
                qq[0, 2] * qq[1, 0] * qq[2, 1] - qq[0, 2] * qq[2, 2] * qq[2, 0] -
                qq[0, 0] * qq[1, 2] * qq[2, 1] - qq[0, 1] * qq[1, 0] * qq[2, 2];
            return det;
        }

        private void рассчитатьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (comboBox1.SelectedIndex == 0)
                {
                    MNK();
                }
            }
            catch
            {
                MessageBox.Show("Что-то пошло не так", "Ошибка!");
            }
        }

        private void рандомнаяГенерацияToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Cl();

            Random rnd = new Random();
            for (int i = 0; i < 10; ++i)
            {
                int value1 = rnd.Next(0, 10);
                int value2 = rnd.Next(0, 10);
                Point xy = new Point(value1, value2);
                XY.Add(xy);
            }

            foreach (var p in XY)
            {
                dataGridView1.Rows.Add(p.x, p.y);
                chart1.Series[1].Points.AddXY(p.x, p.y);
            }
        }

        private void button3_Click(object sender, EventArgs e) // Добавление в DataGridView точек
        {
            try
            {
                if (textBox5.Text == "" || textBox6.Text == "")
                {
                    MessageBox.Show("Заполните оба поля!", "Ошибка!");
                }
                else
                {
                    Point xy = new Point(Convert.ToDouble(textBox5.Text), Convert.ToDouble(textBox6.Text));
                    XY.Add(xy);
                    dataGridView1.Rows.Add(xy.x, xy.y);
                    textBox5.Text = "";
                    textBox6.Text = "";
                    chart1.Series[1].Points.AddXY(xy.x, xy.y);
                }
            }
            catch
            {
                MessageBox.Show("Введите число!", "Ошибка!");
            }
        }

        private void button4_Click(object sender, EventArgs e) // Удаление таблицы
        {
            if (dataGridView1.Rows.Count > 0)
            {
                Cl();
            }
            else
            {
                MessageBox.Show("Таблица пустая!", "Ошибка");
            }
        }
        void Cl()
        {
            dataGridView1.Rows.Clear();
            XY.Clear();
            chart1.Series[0].Points.Clear();
            chart1.Series[1].Points.Clear();
            chart1.Series[2].Points.Clear();
            chart1.Update();
            dataGridView1.Update();
            label1.Text = "";
            label2.Text = "";
        }
    }
}