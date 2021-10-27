using System;
using Aspose.Pdf;
using System.IO;
using System.Windows;
using AutoOrderAPP.Clients;
using System.Collections.Generic;

namespace AutoOrderAPP.Operations
{
    public class ConvertToXls
    {
        public DirectoryInfo directory { get; set; }
        public string FolderPath { get; set; }

        public ConvertToXls(string path)
        {
            FolderPath = path;
        }

        public ConvertToXls()
        {
        }

        Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document();
        ExcelSaveOptions options = new ExcelSaveOptions();

        public FileInfo[] GetAllFiles(string source,string extension)
        {
            directory = new DirectoryInfo(FolderPath + source); 
            try
            {
                FileInfo[] Files = directory.GetFiles(extension);
                return Files;
            }
            finally
            {
                
            }
            
        }

        public void ConvertXls(List<Client> markets)
        {
            
            foreach (var market in markets)
            {
                
                foreach (var filename in GetAllFiles(market.FolderName,"*.pdf"))
                {
                    pdfDocument = new Aspose.Pdf.Document(filename.ToString());
                    options.Format = ExcelSaveOptions.ExcelFormat.XLSX;
                    try
                    {
                        pdfDocument.Save(filename.ToString() + ".xlsx", options);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }

                    GC.Collect();
                }       
            }  
        }
    }
}
