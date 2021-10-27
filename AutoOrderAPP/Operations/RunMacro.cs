using AutoOrderAPP.Clients;
using Scripting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using ClosedXML.Excel;
using File = System.IO.File;

namespace AutoOrderAPP.Operations
{
    public class RunMacro
    {
        public int MyProperty { get; set; }
        public string MacroName { get; set; }
        public string Orderno { get; set; }
        public XLWorkbook WrkBook { get; set; }

        

        public RunMacro()
        {
            MacroName ="macro.txt";
            if (!Directory.Exists(MacroName))
            {
                MacroName = System.IO.File.ReadAllText(MacroName).Trim();
            }
        }

        public void AutoRunMacro(FileInfo [] files,string txtpath)
        {

            foreach (var file in files)
            {
                Excel.Application ExcelApp = new Excel.Application();
                ExcelApp.AutomationSecurity =MsoAutomationSecurity.msoAutomationSecurityForceDisable;
                Excel.Workbook ExcelWorkBook = ExcelApp.Workbooks.Open(file.FullName);
                
                try
                {
                    ExcelApp.Run(MacroName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Unable to Run Macro: " + MacroName + " Exception: " + ex.Message);
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

        public void  OrderNoRecord(string xlspath,string txtpath)
        {
            WrkBook = new XLWorkbook(xlspath);
            var ws1 = WrkBook.Worksheet("Main");
            var data = ws1.Cell("A4").GetValue<string>();
            File.WriteAllTextAsync(txtpath, data);


        }
    }
}
