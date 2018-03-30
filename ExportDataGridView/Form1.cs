using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExportDataGridView
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        struct DataParameter
        {
            public List<String> oneColumn;
            public List<String> twoColumn;
            public List<String> threeColumn;
            public String FileName { get; set; }
        }

        DataParameter _inputParameter;


        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            List<String> one = ((DataParameter)e.Argument).oneColumn;
            List<String> two = ((DataParameter)e.Argument).twoColumn;
            List<String> three = ((DataParameter)e.Argument).threeColumn;
            String FileName = ((DataParameter)e.Argument).FileName;

            Excel.Application excel = new Excel.Application();
            Excel.Workbook wb = excel.Workbooks.Add(Excel.XlSheetType.xlWorksheet);
            Excel.Worksheet ws = (Excel.Worksheet)excel.ActiveSheet;
            excel.Visible = false;
            int index = 1;
            int progress = one.Count;
            //Колонки
            ws.Cells[1, 1] = "Первый)";
            ws.Cells[1, 2] = "Второй=)";
            ws.Cells[1, 3] = "Третий";

            /*
            ((Excel.Range)sheet.Columns).ColumnWidth = 15;
            
            //жирность
            (sheet.Cells[1, 1] as Excel.Range).Font.Bold = true;
            
            //размер шрифта
            (sheet.Cells[1, 1] as Excel.Range).Font.Size = 16;
            
            //название шрифта
            (sheet.Cells[1, 1] as Excel.Range).Font.Name = "Times New Roman";
            
            //стиль границы
            ws.get_Range("B2", "C3").Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            ws.get_Range("B2", "C3").Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            ws.get_Range("B2", "C3").Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            ws.get_Range("B2", "C3").Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            
            //толщина границы
            ws.get_Range("B2", "C3").Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlMedium;
            
            //выравнивание по горизонтали
            (ws.Cells[1, 1] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            
            //выравнивание по вертикали
            (ws.Cells[1, 1] as Excel.Range).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            
            //объединение ячеек
            Excel.Range oRange;
            oRange = ws.Range[ws.Cells[4, 2], ws.Cells[9, 5 + 4]];
            oRange.Merge(Type.Missing);
            */

            //Жирность
            (ws.Cells[1, 1] as Excel.Range).Font.Bold = true;
            (ws.Cells[1, 2] as Excel.Range).Font.Bold = true;
            (ws.Cells[1, 3] as Excel.Range).Font.Bold = true;

            Excel.Range oRange;
            oRange = ws.Range[ws.Cells[4, 2], ws.Cells[9, 5 + 4]];
            oRange.Merge(Type.Missing);
            (ws.Cells[4, 2] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            (ws.Cells[4, 2] as Excel.Range).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            (ws.Cells[4, 2] as Excel.Range).Font.Size = 26;
            (ws.Cells[4, 2] as Excel.Range).Font.Name = "Times New Roman";
            ws.get_Range("B2", "C3").Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            ws.get_Range("B2", "C3").Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            ws.get_Range("B2", "C3").Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            ws.get_Range("B2", "C3").Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            ws.get_Range("B2", "C3").Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlMedium;
            
            for (int i = 0; i < one.Count; i++)
            {
                if (!backgroundWorker1.CancellationPending)
                {
                    backgroundWorker1.ReportProgress(index++ * 100 / progress);
                    ws.Cells[index, 1] = one[i];
                    ws.Cells[index, 2] = two[i];
                    ws.Cells[index, 3] = three[i];
                }
            }

            ws.Columns[1].AutoFit();
            ws.Columns[2].AutoFit();
            ws.Columns[3].AutoFit();
            //Сохранение файла
            ws.SaveAs(FileName, Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, true, false, Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);
            excel.Quit();
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;
            lblStatus.Text = String.Format("Процесс выполнения...{0}%", e.ProgressPercentage);
            progressBar1.Update();
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error == null)
            {
                Thread.Sleep(100);
                lblStatus.Text = "Ваши данные успешно экспортированы.";
            }
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            if (backgroundWorker1.IsBusy)
                return;
            using (SaveFileDialog sfd = new SaveFileDialog() { Filter = "Excel Workbook|*.xls" })
            {
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    _inputParameter.FileName = sfd.FileName;
                    _inputParameter.oneColumn = new List<string>();
                    _inputParameter.twoColumn = new List<string>();
                    _inputParameter.threeColumn = new List<string>();

                    for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                    {
                        if (dataGridView1.Rows[i].Cells[0].Value != null)
                            _inputParameter.oneColumn.Insert(i, dataGridView1.Rows[i].Cells[0].Value.ToString());
                        else
                            _inputParameter.oneColumn.Insert(i, "");

                        if (dataGridView1.Rows[i].Cells[1].Value != null)
                            _inputParameter.twoColumn.Insert(i, dataGridView1.Rows[i].Cells[1].Value.ToString());
                        else
                            _inputParameter.twoColumn.Insert(i, "");

                        if (dataGridView1.Rows[i].Cells[2].Value != null)
                            _inputParameter.threeColumn.Insert(i, dataGridView1.Rows[i].Cells[2].Value.ToString());
                        else
                            _inputParameter.threeColumn.Insert(i, "");
                    }
                    progressBar1.Minimum = 0;
                    progressBar1.Value = 0;
                    backgroundWorker1.RunWorkerAsync(_inputParameter);
                }
            }
        }
    }
}
