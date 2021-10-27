using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AutoMacro
{
    public partial class Form1 : Form
    {

        Test test = new Test();
        List<string> filenames = new List<string>();


        public Form1()
        {
            string folderpath = @"C:\\Users\\Xeyyam\\Desktop\\test";
           // InitializeComponent();
            MessageBox.Show("Baslayir");
            

            DirectoryInfo d = new DirectoryInfo(folderpath); //Assuming Test is your Folder

            FileInfo[] Files = d.GetFiles("*.xlsx"); //Getting Text files            

            foreach (var filename in Files)
            {
                test.AutoRunMacro(filename.ToString());
              
            }
            GC.Collect();
            MessageBox.Show("Bitdi");

        }
    }
}
