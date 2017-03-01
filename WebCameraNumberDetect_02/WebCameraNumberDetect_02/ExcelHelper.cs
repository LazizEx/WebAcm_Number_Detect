using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace WebCameraNumberDetect_02
{
    class ExcelHelper
    {

        public ExcelHelper()
        {

        }

        public void Create(List<ExcelValues> dict)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            xlWorkSheet.Cells[1, 1] = "Time";
            xlWorkSheet.Cells[1, 2] = "Value";

            for (int i = 0; i < dict.Count; i++)
            {
                xlWorkSheet.Cells[i + 2, 1] = string.Format("{0}:{1}:{2}", dict[i].time.Hours, dict[i].time.Minutes, dict[i].time.Seconds);
                xlWorkSheet.Cells[i + 2, 2] = dict[i].value;
            }
            //xlWorkSheet.Cells[2, 1] = "00:0:0";
            //xlWorkSheet.Cells[2, 2] = "25.8";

            Excel.Range chartRange;

            Excel.ChartObjects xlCharts = (Excel.ChartObjects)xlWorkSheet.ChartObjects(Type.Missing);
            Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(10, 80, 300, 250);
            Excel.Chart chartPage = myChart.Chart;
            string c = (dict.Count + 1).ToString();
            chartRange = xlWorkSheet.get_Range("A1", "b" + c);
            chartPage.SetSourceData(chartRange, misValue);
            chartPage.ChartType = Excel.XlChartType.xlLine;

            string f = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            try
            {
                xlWorkBook.SaveAs(System.IO.Path.Combine(f, @"informations.xls"), Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
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
                //MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
