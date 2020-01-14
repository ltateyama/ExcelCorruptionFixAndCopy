using System;
using System.IO;
using System.Threading;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;


namespace ExcelCorruptionFixAndCopy
{
    class Program
    {
        static void Main(string[] args)
        {
            RepairAndCopyExcelFile(args);
            GC.Collect();
            GC.WaitForPendingFinalizers();

            Console.WriteLine("Verifying file released");
            FileAvail(args[1]);
            Console.WriteLine("Done.");
        }


        static void RepairAndCopyExcelFile(string[] args)
        {
            if (args.Length < 2)
            {
                Console.WriteLine("Missing Required Arguments");
            }
            else
            {
                string strSrcFile = args[0];
                string strDstFile = args[1];

                Console.WriteLine(strSrcFile);
                Console.WriteLine(strDstFile);

                Excel.Application xlApp     = null;
                Excel.Workbooks xlWorkbooks = null;
                Excel.Workbook xlWorkbook   = null;

                try
                {
                    Console.WriteLine("Openning xlApp");
                    xlApp = new Excel.Application()
                    {
                        DisplayAlerts = false
                    };
                    Console.WriteLine("Openning Workbooks");
                    xlWorkbooks = xlApp.Workbooks;

                    Console.WriteLine("Openning Workbook");
                    xlWorkbook  = xlWorkbooks.Open(strSrcFile, CorruptLoad: Excel.XlCorruptLoad.xlRepairFile);
                }
                finally
                {
                    Console.WriteLine("Saving " + strDstFile);
                    xlWorkbook.SaveAs(strDstFile);

                    Console.WriteLine("Closing Workbook");
                    xlWorkbook.Close(0);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkbook);

                    Console.WriteLine("Closing Workbooks");
                    xlWorkbooks.Close();
                    if (xlWorkbooks != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkbooks);

                    Console.WriteLine("Closing xlApp");
                    xlApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
                }
            }
        }
        static void FileAvail (string strFilePath)
        {
            bool bolSuccess = false;
            bool bolSkipFailure = false;
            Int32 intRetryCount = 0;
            Int32 intDelay = 1;
            do
            {
                try
                {
                    using (File.Open(strFilePath, FileMode.Open)) {}
                    bolSuccess = true;
                }
                catch (IOException e)
                {
                    Console.WriteLine("\t\tExel File: " + strFilePath + " Failed to open " + intRetryCount + " time(s).");

                    if (intRetryCount++ < 10)
                    {
                        Thread.Sleep(intDelay * 1000);
                        intDelay = intDelay * 2;
                    }
                    else
                    {
                        bolSkipFailure = true;
                    }
                }
            } while ((!bolSuccess) && (!bolSkipFailure));

        }
    }
}
