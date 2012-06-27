using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;

using log4net;
using MSExcel = Microsoft.Office.Interop.Excel;


namespace JustAProgrammer.PIA.Excel
{
    class Program
    {
        private static ILog log = LogManager.GetLogger(typeof(Program));
        static void Main(string[] args) {

            var xlApp = new MSExcel.ApplicationClass();
            var xlWorkbooks = xlApp.Workbooks;
            xlApp.Visible = false;
            xlApp.DisplayAlerts = false;

            for (int i = 0; i < 500; i++)
            {

                var workbook = Path.Combine(Environment.CurrentDirectory, "us_foreign_assistance.xls");
                //var workbook = "https://explore.data.gov/download/5gah-bvex/XLS";
                var xlWorkBook = xlWorkbooks._Open(workbook, Missing.Value, Missing.Value, Missing.Value,
                                                       Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                                                       Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                                                       Missing.Value);
                

                var xlWorkSheet = (MSExcel._Worksheet)xlWorkBook.Sheets["Notes"];
                xlWorkSheet.Delete();

                xlApp.Visible = false;
                workbook = Path.Combine(Environment.CurrentDirectory, string.Format("Book-{0}.xls", i));
                xlWorkBook.SaveAs(workbook, MSExcel.XlFileFormat.xlWorkbookNormal, Missing.Value, Missing.Value,
                                  false, false, MSExcel.XlSaveAsAccessMode.xlNoChange, Missing.Value, Missing.Value,
                                  Missing.Value, Missing.Value, Missing.Value);
                // Commenting this out will cause memory usage to spike.
                xlWorkBook.Close(false, Missing.Value, false);
                log.InfoFormat("Wrote \"{0}\".", workbook);
                // Commenting this out does not seem to affect memory usage.
                Marshal.ReleaseComObject(xlWorkBook);
            }
            xlApp.DisplayAlerts = true;
            xlApp.Quit();
            Marshal.ReleaseComObject(xlWorkbooks);
            Marshal.ReleaseComObject(xlApp);

            /*
            Console.Write("Press any key to continue . . . ");
            Console.ReadKey(false);
            Console.WriteLine();
             */
        }
    }
}
