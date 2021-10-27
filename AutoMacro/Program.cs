using Microsoft.ReportingServices.ReportProcessing.ReportObjectModel;
using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;




namespace AutoMacro
{
    public class Test
    {


        string macro = "PERSONAL.XLSB!CSH_APP";
       
        
        public void AutoRunMacro(string sourceFile)
        {
            Excel.Application ExcelApp = new Excel.Application();
            ExcelApp.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable;
            Excel.Workbook ExcelWorkBook = ExcelApp.Workbooks.Open(sourceFile);

            
            try
            {
                ExcelApp.Run(macro);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Unable to Run Macro: " + macro + " Exception: " + ex.Message);
            }
            
            ExcelWorkBook.Close(true);
            ExcelApp.Quit();
            //GC.Collect();
            
           if (ExcelWorkBook != null) { System.Runtime.InteropServices.Marshal.ReleaseComObject(ExcelWorkBook); }
           if (ExcelApp != null) { System.Runtime.InteropServices.Marshal.ReleaseComObject(ExcelApp); }

            ExcelApp = null;
            GC.Collect();
        }


    }


    static class Program
    {

       



        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        public static void Main()
        {


            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());

            MessageBox.Show("huuuu");



        }





    }
}
