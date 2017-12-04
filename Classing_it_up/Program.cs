using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using Marshal = System.Runtime.InteropServices;

namespace Classing_it_up
{
    class Program
    {
        static void Main(string[] args)
        {
            //setup log file information
            string logPath = @"\\cfc1afs01\Operations-Analytics\Log_Files\";
            StreamWriter logging = null;
            logging = new StreamWriter(logPath + "ClassyLog.txt");
            logging.WriteLine("Log Started for Classy run at " + System.DateTime.Now);
            Console.WriteLine("Log Started for Classy run at " + System.DateTime.Now);

            //Setup used Ranges variables
            int srcUsedRange = 0;
            int destUsedRange = 0;
            int srcUsedRange2 = 0;
            int destUsedRange2 = 0;
            int destUsedRange3 = 0;
            int destUsedRange4 = 0;
           // int y = 2;
            //string filename = "";
            //string a = "";
            //string b = "";
           // double percentComplete = 0.0;
           // string shift = "";
            logging.WriteLine(System.DateTime.Now + ":\t" + "Success in Declaring Variables srcUsedRange & destUsedRange.");
            Console.WriteLine(System.DateTime.Now + ":\t" + "Success in Declaring Variables srcUsedRange & destUsedRange.");

            //setup Excel Application
            Excel.Application App = null;
            App = new Excel.Application();
            logging.WriteLine(System.DateTime.Now + ":\t" + "Success in setting up application Reference for Excel Application.");
            Console.WriteLine(System.DateTime.Now + ":\t" + "Success in setting up application Reference for Excel Application.");

            //Setup path variable
            string srcData = @"\\cfc1afs01\Operations-Analytics\Projects\ClassingTool\RAW_CLASS_DATA.xlsx";
            string ClassingPlan = @"\\CFC1AFS01\Operations-Analytics\Projects\ClassingTool\ClassingPlan v6(MCRC).xlsx";
            string ClassingTool = @"\\CFC1AFS01\Operations-Analytics\Projects\ClassingTool\Classing Tool v3(MCRC).xlsx";
            string path2 = @"\\CFC1AFS01\Operations-Analytics\Projects\ClassingTool\";

            //create soure and destination files
            Excel.Workbooks Workbooks = null;
            Excel.Workbook srcDataWorkbook = null;
            Excel.Worksheet srcDataWorksheet = null;
            Excel.Worksheet srcDataWorksheet2 = null;
            Excel.Workbook destCLSTLNWorkbook = null;
            Excel.Worksheet destCLSTLNWorksheet = null;
            Excel.Worksheet destCLSTLNWorksheet2 = null;
            Excel.Workbook destCLSPLNWorkbook = null;
            Excel.Worksheet destCLSPLNWorksheet = null;
            Excel.Worksheet destCLSPLNWorksheet2 = null;

            //Assign the source and destination files
            Workbooks = App.Workbooks;
            srcDataWorkbook = Workbooks.Open(srcData, false, false);
            srcDataWorksheet = srcDataWorkbook.Worksheets.Item["ItemDimWMax"];
            srcDataWorksheet2 = srcDataWorkbook.Worksheets.Item["6WksReceiving"];
            destCLSTLNWorkbook = Workbooks.Open(ClassingTool, false, false);
            destCLSPLNWorkbook = Workbooks.Open(ClassingPlan, false, false);
            destCLSTLNWorksheet = destCLSTLNWorkbook.Worksheets.Item["Inv"];
            destCLSTLNWorksheet2 = destCLSTLNWorkbook.Worksheets.Item["Receiving"];
            destCLSPLNWorksheet = destCLSPLNWorkbook.Worksheets.Item["Full Item Catalog"];
            destCLSPLNWorksheet2 = destCLSPLNWorkbook.Worksheets.Item["Current Inventory"];
            logging.WriteLine(System.DateTime.Now + ":\t" + "Success in setting up src & dest Workbooks & Worksheets.");
            Console.WriteLine(System.DateTime.Now + ":\t" + "Success in setting up src & dest Workbooks & Worksheets.");

            //setup Excel enviroment, Stop calcuation before copy
            App.Application.DisplayAlerts = false;
            App.Application.Visible = true;
            App.Application.Calculation = Excel.XlCalculation.xlCalculationManual;
            logging.WriteLine(System.DateTime.Now + ":\t" + "Successfully turned off application alerts and set workbook to manual calculation.");
            Console.WriteLine(System.DateTime.Now + ":\t" + "Successfully turned off application alerts and set workbook to manual calculation.");

            //Gather the size of the range to copy
            srcUsedRange = srcDataWorksheet.UsedRange.Rows.Count;
            srcUsedRange2 = srcDataWorksheet2.UsedRange.Rows.Count;
            destUsedRange = destCLSTLNWorksheet.UsedRange.Rows.Count;
            destUsedRange2 = destCLSTLNWorksheet2.UsedRange.Rows.Count;
            destUsedRange3 = destCLSPLNWorksheet.UsedRange.Rows.Count;
            destUsedRange4 = destCLSPLNWorksheet2.UsedRange.Rows.Count;

            //clear the destination range so we do not get ghosting data
            Excel.Range r1 = destCLSTLNWorksheet.Cells[1, 1];
            Excel.Range r2 = destCLSTLNWorksheet.Cells[destUsedRange, 26];
            Excel.Range destRange = destCLSTLNWorksheet.Range[r1, r2];
            destRange.Value = "";
            Excel.Range r3 = destCLSTLNWorksheet2.Cells[1, 1];
            Excel.Range r4 = destCLSTLNWorksheet2.Cells[destUsedRange2, 13];
            Excel.Range destRange2 = destCLSTLNWorksheet2.Range[r3, r4];
            destRange.Value = "";
            Excel.Range r5 = destCLSPLNWorksheet.Cells[1, 1];
            Excel.Range r6 = destCLSPLNWorksheet.Cells[srcUsedRange, 43];
            Excel.Range destRange3 = null;
            logging.WriteLine(System.DateTime.Now + ":\t" + "Successfully cleared destionation range.");
            Console.WriteLine(System.DateTime.Now + ":\t" + "Successfully cleared destionation range.");

            //setup destination range varabiles
            r2 = destCLSTLNWorksheet.Cells[srcUsedRange, 26];
            destRange = destCLSTLNWorksheet.Range[r1, r2];
            r4 = destCLSTLNWorksheet2.Cells[srcUsedRange2, 13];
            destRange2 = destCLSTLNWorksheet2.Range[r3, r4];
            

            //Copy source to it's destination
            destRange.Value = srcDataWorksheet.UsedRange.Value;
            destRange2.Value = srcDataWorksheet2.UsedRange.Value;
            r2 = destCLSTLNWorksheet.Cells[srcUsedRange, 43];
            destRange = destCLSTLNWorksheet.Range[r1, r2];
           // destRange3 = destCLSPLNWorksheet.Range[r5, r6];
           // destRange3.Value = destCLSTLNWorksheet.UsedRange.Value;

            //clean up Function assignment
            r1 = destCLSTLNWorksheet.Cells[27, 3];
            r2 = destCLSTLNWorksheet.Cells[destUsedRange, 17];
            Excel.Range srcfuctcopy = destCLSTLNWorksheet.Range["AA3:AQ3"];
            Excel.Range fillRange = destCLSTLNWorksheet.Range[r1, r2];
            //srcfuctcopy.AutoFill(destCLSTLNWorksheet.Range[r1,r2]);
           

            //Setup sheet Variables
            destCLSTLNWorkbook.Worksheets.Item["Tolerances"].cells[2, 2].value = System.DateTime.Today;

            //Restore calcuation
            App.Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic;

            //close the worbooks saving destination
            destCLSPLNWorkbook.SaveAs(path2 + "ClassingPlan v6.xlsx");
            destCLSTLNWorkbook.SaveAs(path2 + "Classing Tool v3.xlsx");
            destCLSTLNWorkbook.Close(false);
            destCLSPLNWorkbook.Close(false);
            srcDataWorkbook.Close(false);
            logging.WriteLine(System.DateTime.Now + ":\t" + "Successfully saved Destination Workbook and closed sourceWorkbook.");

            //Null out the remaining Excel variables
            destCLSTLNWorkbook = null;
            destCLSTLNWorksheet = null;
            srcDataWorkbook = null;
            srcDataWorksheet = null;

            //quit the excel application
            App.Quit();
            App = null;
            logging.WriteLine(System.DateTime.Now + ":\t" + "Successfully closed references to the Excel Application.");

            //clear com object references
            if (destCLSTLNWorkbook != null)
            { Marshal.Marshal.ReleaseComObject(destCLSTLNWorkbook); }
            if (destCLSTLNWorksheet != null)
            { Marshal.Marshal.ReleaseComObject(destCLSTLNWorksheet); }
            if (srcDataWorkbook != null)
            { Marshal.Marshal.ReleaseComObject(srcDataWorkbook); }
            if (srcDataWorksheet != null)
            { Marshal.Marshal.ReleaseComObject(srcDataWorksheet); }
            if (Workbooks != null)
            { Marshal.Marshal.ReleaseComObject(Workbooks); }
            if (App != null)
            { Marshal.Marshal.ReleaseComObject(App); }

            //Yell Success!
            Console.WriteLine("HORRAY Execution Completed!");
            logging.WriteLine(System.DateTime.Now + ":\t" + "Execution Completed");
            //close reference to streamwriter
            logging.Close();
            logging = null;

        }
    }
}
