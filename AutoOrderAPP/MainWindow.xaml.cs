using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using AutoOrderAPP.Clients;
using AutoOrderAPP.Clients.Bazarstore;
using AutoOrderAPP.Operations;
using AutoOrderAPP.Orders.AutoOrderAPP.FileService;
using AutoOrderAPP.UI;
using Limilabs.Client.IMAP;
using Limilabs.Mail;
using Limilabs.Mail.MIME;
using Imap = Limilabs.Client.IMAP.Imap;


namespace AutoOrderAPP
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        string dateToDay = DateTime.Now.ToString("dd.MM.yyyy"); //Date Format "09.19.2021"
        string weekOfDay = DateTime.Now.ToString("dddd");       //name of weekday exp:("Monday,Thursday")

        string MailFolderPath = @"D:\C# Projects\AutoOrderAPP\AutoOrderAPP\UI\DataFill\mail_folder.txt";

        Client client = new Client();

        StatusLoginLogout status = new StatusLoginLogout();
        
        FileService fileService = new FileService();

        ConvertToXls converter = new ConvertToXls();

        RunMacro runMacro = new RunMacro();

        List<Client> RahatTodayFolders, BazarstoreTodayFolders = new List<Client>();

        

        



        public MainWindow()
        {   
            
            InitializeComponent();
            if (weekOfDay != "Sunday")
            {
                
              // BazarstoreTodayFolders = client.GetAllFolderName(Bazarstore.RouteFolder, weekOfDay);
               RahatTodayFolders = client.GetAllFolderName(Rahat.RouteFolder,weekOfDay);

              // client.CreateRootFolders(dateToDay, Bazarstore.SpecialName, Bazarstore.RouteFolder, BazarstoreTodayFolders);
               client.CreateRootFolders(dateToDay,Rahat.SpecialName,Rahat.RouteFolder,RahatTodayFolders);

                
                if (File.Exists(MailFolderPath))
                {
                    string[]  temp = new string[] { };
                    temp = File.ReadAllLines(MailFolderPath);
                    for (int i = 0; i < temp.Length; i++)
                    {   
                        if (temp[i]!="")
                        folder_name.Items.Add(temp[i].Trim());
                    }
                }

                foreach (var market in RahatTodayFolders)
                { 
                    mrkt_names.Items.Add(market.MarketName.Trim());
                }


            }


            

        }


        
        private void btnConvert_Click(object sender, RoutedEventArgs e)
        {  
            if (folder_name.SelectedIndex != -1)
            {
                btnConvert.IsEnabled = false;
                converter = new ConvertToXls(folder_name.Text.Trim() + "\\" + dateToDay + "\\" + Rahat.SpecialName + "\\");
                converter.ConvertXls(RahatTodayFolders);
                GC.Collect();
                txtbox_info.Text += "\n All .pdf files was converted to .xls tpye of file";
                btnRunMacro.IsEnabled = true;
            }
            
        }

        private void btnDownload_Click(object sender, RoutedEventArgs e)
        {
            
            if (folder_name.SelectedIndex != -1)
            {
                string temp = folder_name.SelectedItem.ToString().Trim();
                using (status.imap)
                {
                    btnDownload.IsEnabled = false;
                    txtbox_info.Text += "\n Yuklenme davam edir...";
                    int count =fileService.DownloadFromMail(lgn_txt.Text,txtpwd.Password, dateToDay, new Client(Rahat.RouteFolder), RahatTodayFolders, temp);
                    txtbox_info.Text +=count.ToString()+" sayda fayl yuklendi \n Yuklenme Bitdi!";
                    btnConvert.IsEnabled = true;

                }
              
                
            }
            else
            {
                MessageBox.Show("Zəhmət olmasa Faylları yükləmək istədiyiniz qovluğu seçin!");
            }

            
            
            
        }

       

        private void btnRunMacro_Click(object sender, RoutedEventArgs e)
        {
            if (folder_name.SelectedIndex != -1) 
            {
                converter = new ConvertToXls(folder_name.Text.Trim() + "\\" + dateToDay + "\\" + Rahat.SpecialName + "\\");
                foreach (var folder in RahatTodayFolders)
                {
                    runMacro.AutoRunMacro(converter.GetAllFiles(folder.FolderName, "*.xlsx"),converter.FolderPath);
                }
                btnRunMacro.IsEnabled = false;
                GC.Collect();
                //upload button enabled=true
                txtbox_info.Text += "\n All .xls files was modify  with specify Macro!!!";
            }
           
        }



        //mail log and out  buttons click functions
        private void btn_exitmail_click(object sender, RoutedEventArgs e)
        {  
            if (status.LogConfirm)
            {
                status.LogOutMail();
                btn_login.IsEnabled = true;
                btnExitMail.IsEnabled = false;
            }
            else
            {
                MessageBox.Show("Please login before this operation!");
            }
            
        }

        private void btn_login_click(object sender, RoutedEventArgs e)
        {



            if (lgn_txt.Text.Length >= 10 && txtpwd.Password.Length >= 8)
            {

                status.LoginMail("yandex.ru", lgn_txt.Text, txtpwd.Password);
                lgn_txt.Text = null;
                txtpwd.Password = null;
                txtbox_info.Text += "\n" + status.Result;
                if (status.LogConfirm)
                {
                    btn_login.IsEnabled = false;
                    btnDownload.IsEnabled = true;
                    btnExitMail.IsEnabled = true;
                }


                
            }


        }

    }
}
