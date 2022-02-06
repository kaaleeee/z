using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace trpo2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private readonly string excelSavePath = @"C:\Users\miroo\OneDrive\Рабочий стол\tabl";

        private void button6_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                if (comboBox1.Text == Convert.ToString(dataGridView1[2, i].Value))
                {
                    dataGridView1[4, i].Value = (Convert.ToInt32(dataGridView1[4, i].Value) / 100 * Convert.ToInt32(textBox1.Text)) + Convert.ToInt32(dataGridView1[4, i].Value);

                }
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            double gg = 0;
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
               
                if (comboBox2.Text == Convert.ToString(dataGridView1[3, i].Value))
                {
                    gg += 1;

                }
                textBox2.Text = gg.ToString();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            double sum = 99999999999;
            for (int i = 0; i < dataGridView1.RowCount - 1; i++)
            {
                if (dataGridView1.RowCount < sum)
                    sum = Convert.ToDouble(dataGridView1[4, i].Value);
            }
            textBox4.Text = sum.ToString();
        }

        private void button4_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count > 1)
            {
                dataGridView1.Rows.RemoveAt(dataGridView1.CurrentRow.Index);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            double sum = 0;
            for (int i = 0; i < dataGridView1.RowCount - 1; i++)
            {
                sum += Convert.ToDouble(dataGridView1[4, i].Value);
            }
            textBox3.Text = sum.ToString();
        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application appExcel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook workBookExcel = appExcel.Workbooks.Add(Type.Missing);
            Microsoft.Office.Interop.Excel.Worksheet worksheetExcel = null;
            appExcel.Visible = true;
            worksheetExcel = workBookExcel.Sheets[1];
            worksheetExcel = workBookExcel.ActiveSheet;
            worksheetExcel.Name = "Таблица 1";
            //Копируем заголовки
            for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
            {
                worksheetExcel.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
            }
            try
            {
                for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                {
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        worksheetExcel.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                    }
                }
            }
            catch
            {
                MessageBox.Show("Таблица сохранена", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            //Сохраняем
            workBookExcel.SaveAs(excelSavePath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            appExcel.Quit();

        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            {
                string str;
                int rCnt;
                int cCnt;
                OpenFileDialog opf = new OpenFileDialog();
                opf.Filter = "Excel (*.xlsx)|*.xlsx";
                opf.ShowDialog();
                System.Data.DataTable tb = new System.Data.DataTable();
                string filename = opf.FileName;
                Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel._Workbook ExcelWorkBook;
                Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet;
                Microsoft.Office.Interop.Excel.Range ExcelRange;
                ExcelWorkBook = ExcelApp.Workbooks.Open(filename, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false,
                    false, 0, true, 1, 0);
                ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);
                ExcelRange = ExcelWorkSheet.UsedRange;
                for (rCnt = 1; rCnt <= ExcelRange.Rows.Count; rCnt++)
                {
                    dataGridView1.Rows.Add(1);
                    for (cCnt = 1; cCnt <= 6; cCnt++)
                    {
                        str = Convert.ToString((ExcelRange.Cells[rCnt, cCnt] as Microsoft.Office.Interop.Excel.Range).Value2);
                        dataGridView1.Rows[rCnt - 1].Cells[cCnt - 1].Value = str;
                    }
                }
                ExcelWorkBook.Close(true, null, null);
                ExcelApp.Quit();
                releaseObject(ExcelWorkSheet);
                releaseObject(ExcelWorkBook);
                releaseObject(ExcelApp);
            }

        }
    }
}
