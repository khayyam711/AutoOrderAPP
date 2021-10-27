using Aspose.Pdf;
using System;
using System.IO;
using System.Windows;



namespace Test

{
    public sealed class InputManager : System.Windows.Threading.DispatcherObject
    {
        public InputManager()
        {
        }

        public static object Current { get; internal set; }
    }

    

    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window



    {


        public MainWindow()
        {


            MessageBox.Show("PDF-> XLS Start!!!");
            ConvertXls("C:\\Users\\Xeyyam\\Desktop\\test");
            MessageBox.Show("PDF-> XLS END!!!");
            Close();
                
            //InitializeComponent();
            

            

        }

        Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document();
        ExcelSaveOptions options = new ExcelSaveOptions();
        DirectoryInfo d = new DirectoryInfo("C:\\Users\\Xeyyam\\Desktop\\test");
        public void ConvertXls(string folderpath)
        {

            d = new DirectoryInfo(folderpath); 
            FileInfo[] Files = d.GetFiles("*.pdf");
            foreach (var filename in Files)
            {
                pdfDocument = new Aspose.Pdf.Document(filename.ToString());
                options.Format = ExcelSaveOptions.ExcelFormat.XLSX;
                try
                {
                    pdfDocument.Save(filename.ToString() + ".xlsx", options);
                }
                catch(Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                
               // GC.Collect();
            }
            
        }






    }
}
