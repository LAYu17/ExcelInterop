using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace KursachExcelInterop
{
    public partial class MainF : Form
    {
        private Excel.Application _excel;
        private Excel._Workbook _book;
        private Excel._Worksheet _sheet;
        private bool _isClosingExcel;
        private List<PointF> al;
        private Graphics gr;
        private ImageData[] ImData;
        private int count1 = 0;
        private int count2 = 0;
        private int count3 = 0;

        public MainF() => InitializeComponent();

        private void btnOpen_Click(object sender, EventArgs e)
        {
            _excel = new Excel.Application { Visible = true };
            _book = _excel.Workbooks.Open(Path.GetDirectoryName(Application.ExecutablePath) + "\\DISKRF.xlsx");
            _excel.WorkbookBeforeClose += _excel_WorkbookBeforeClose;
            _sheet = _excel.Sheets[1];
            _sheet.Activate();
            button1.Enabled = true;
        }

        private void _excel_WorkbookBeforeClose(Excel.Workbook wb, ref bool cancel)
        {
            cancel = _isClosingExcel;
            _excel.Visible = false;
            ReadData(gr);
        }

        private void ReadData(Graphics g)
        {
            panel1.Invoke(new MethodInvoker(() => panel1.Refresh()));
            MyEllipse m1 = new MyEllipse();
            MyRectangle r1 = new MyRectangle();
            MyPie p1 = new MyPie();
            var data1 = ReadExcelArray2Dim(_sheet, 30, 2, 2, 1);
            var data2 = ReadExcelArray2Dim(_sheet, 30, 2, 2, 4);
            var data3 = ReadExcelArray2Dim(_sheet, 30, 2, 2, 7);
            var data5 = ReadExcelArray2Dim(_sheet, 6, 2, 2, 39);
            ImData = ReadExcelImageData(_sheet, 30, 2, 10, 42);
            _isClosingExcel = true;
            CloseExcel();
            _isClosingExcel = false;
            СalculationEl(data1, g, m1, Color.Red);
            СalculationRect(data2, g, r1);
            СalculationPie(data3, g, p1);
            СalculationEl1(ImData, g, m1, Color.Blue);
            CalculateLines(data5, g);
        }

        private static ImageData[] ReadExcelImageData(Excel._Worksheet sheet, int rowCount, int startRow, int startColumnPoints, int columnResults)
        {
            var pts = ReadExcelArray2Dim(sheet, rowCount, 2, startRow, startColumnPoints);
            var rs = ReadExcelCheckResults(sheet, rowCount, startRow, columnResults);
            var res = new ImageData[rowCount];
            for (var y = 0; y < rowCount; y++)
                res[y] = new ImageData(new PointF((float)pts[y, 0], (float)pts[y, 1]), rs[y]);
            return res;
        }

        private static int[] ReadExcelCheckResults(Excel._Worksheet sheet, int rowCount, int startRow, int column)
        {
            var ret = new int[rowCount];
            for (var y = 0; y < rowCount; y++)
                ret[y] = int.Parse(sheet.Cells[column][y + startRow].Text);
            return ret;
        }

        private static double[,] ReadExcelArray2Dim(Excel._Worksheet sheet, int rowCount, int columnCount, int startRow, int startColumn)
        {
            var ret = new double[rowCount, columnCount];
            for (var y = 0; y < rowCount; y++)
                for (var x = 0; x < columnCount; x++)
                    ret[y, x] = double.Parse(sheet.Cells[x + startColumn][y + startRow].Text,CultureInfo.InvariantCulture);
            return ret;
        }

        private void MainF_FormClosing(object sender, FormClosingEventArgs e)
        {
            _isClosingExcel = true;
            CloseExcel();
        }

        private void CloseExcel()
        {
            _excel?.Quit();
            _excel = null;
            _book = null;
            _sheet = null;
        }

        private void СalculationEl(double[,] dat, Graphics g, MyEllipse m1, Color col)
        {
            for (int i = 0; i < 30; i++)
            {
                m1 = new MyEllipse(dat[i, 0] - 3, dat[i, 1] - 3, 6, 6, Color.Black, col);
                m1.Draw(g);
            }
        }

        private void СalculationEl1(ImageData[] dat, Graphics g, MyEllipse m1, Color col)
        {
            for (int i = 0; i < dat.Length; i++)
            {
                m1 = new MyEllipse(dat[i].Point.X - 3, dat[i].Point.Y - 3, 6, 6, Color.Black, col);
                m1.Draw(g);
            }
        }

        private void СalculationRect(double[,] dat, Graphics g, MyRectangle m1)
        {
            for (int i = 0; i < 30; i++)
            {
                m1 = new MyRectangle(dat[i, 0] - 3, dat[i, 1] - 3, 6, 6, Color.Black, Color.Green);
                m1.Draw(g);
            }
        }

        private void СalculationPie(double[,] dat, Graphics g, MyPie m1)
        {
            for (int i = 0; i < 30; i++)
            {
                m1 = new MyPie((float)dat[i, 0] - 3, (float)dat[i, 1] - 3, 6, 6, Color.Black, Color.Yellow);
                m1.Draw(g);
            }
        }

        private void CalculateLines(double[,] dat, Graphics g)
        {
            Pen p = new Pen(Color.Brown);
            Pen p22 = new Pen(Color.DarkMagenta);
            Pen p33 = new Pen(Color.PaleVioletRed);
            var p1 = new PointF((float)dat[0, 0], (float)dat[0, 1]);
            var p2 = new PointF((float)dat[1, 0], (float)dat[1, 1]);
            g.DrawLine(p, p1, p2);
            var p3 = new PointF((float)dat[2, 0], (float)dat[2, 1]);
            var p4 = new PointF((float)dat[3, 0], (float)dat[3, 1]);
            g.DrawLine(p22, p3, p4);
            var p5 = new PointF((float)dat[4, 0], (float)dat[4, 1]);
            var p6 = new PointF((float)dat[5, 0], (float)dat[5, 1]);
            g.DrawLine(p33, p5, p6);
        }

        private void MainF_Load(object sender, EventArgs e)
        {
            button1.Enabled = false;
            btnOpen.Enabled = false;
            gr = panel1.CreateGraphics();
            al = new List<PointF>();
            pictureBox1.Image = Properties.Resources._111;
            pictureBox2.Image = Properties.Resources.materEcs;
            pictureBox3.Image = Properties.Resources._2nd;
            pictureBox4.Image = Properties.Resources.Без_имени;
        }
        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            
            MessageBox.Show($"Программа предназначена для распознавания образов методом дискриминантной функции.\r1)Открываем документ кнопкой <<Oткрыть Excel>>.\r2)Откройте вкладку Файл, нажмите кнопку Параметры и выберите категорию Надстройки.\r3)В раскрывающемся списке Управление выберите пункт Надстройки Excel и нажмите кнопку Перейти.\r4)В диалоговом окне Надстройки установите флажок Пакет анализа, а затем нажмите кнопку ОК.\r5)Откройте вкладку Данные, выберите Анализ данных, далее Генерация случайных чисел и жмем ОК.\r6)Теперь заполняем образы (нормальное распределение) и материал экзамена (равномерное распределение) по вашим предпочтениям.\r\rПосле прочтения данной инструкции, поставьте галочку под кнопкой <<Oткрыть Excel>> ");
        }
        private void panel1_MouseClick(object sender, MouseEventArgs e)
        {
            for (var i = 0; i < ImData.Length; i++)
            {
                var t = ImData[i];
                
                if (!t.Triggerctangle.Contains(e.Location)) continue;
                MessageBox.Show($"Принадлежит образу {t.CheckRes}, индекс точки {i}");
                break;
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
                btnOpen.Enabled = true;
            else
            {
                btnOpen.Enabled = false;
            }
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            for (var i = 0; i < ImData.Length; i++)
            {
                var t = ImData[i];
                if (t.CheckRes == 1)
                    count1++;
                if (t.CheckRes == 2)
                    count2++;
                if (t.CheckRes == 3)
                    count3++;
            }

            MessageBox.Show($" Кол-во точек: {count1} принадлежат образу 1\r Кол-во точек: {count2} принадлежат образу 2\r Кол-во точек: {count3} принадлежат образу 3");
        }
        private void button2_Click(object sender, EventArgs e)
        {
            _excel = new Excel.Application { Visible = true };
            _book = _excel.Workbooks.Open(Path.GetDirectoryName(Application.ExecutablePath) + "\\diskrFOR2.xlsx");
            _excel.WorkbookBeforeClose += _excel_WorkbookBeforeClose;
            _sheet = _excel.Sheets[1];
            _sheet.Activate();
            button1.Enabled = true;
        }
    }
}