using System;
using Imap = Limilabs.Client.IMAP.Imap;
using System.Windows;
using MailKit.Net.Imap;
using Aspose.Email.Clients;
using AutoOrderAPP.Orders.AutoOrderAPP.FileService;

namespace AutoOrderAPP.UI
{
    public class StatusLoginLogout
    {
        public Imap imap { get; set; }
        public string Result { get; set; }
        public bool LogConfirm = false;

        FileService fileservice_ob = new FileService();


        public void LoginMail(string url,string login,string key)
        {
 
                using (this.imap = new Imap())
                {
                    this.imap.ConnectSSL("imap."+url, 993);
                    try
                    {
                        this.imap.UseBestLogin(login, key);
                        
                       
                        
                        this.LogConfirm = true;
                        this.Result="Connected!!!";
                        


                    }
                    catch 
                    {
                        
                        MessageBox.Show("Login or password is incorrect, please try again!");
                    }

                }imap.Close(false);
            }



        //public void LoginManual()
        //{
        //    ImapClient client = new ImapClient("imap.domain.com", 993, "user@domain.com", "pwd");

        //    // Set the security mode to implicit
        //    client.SecurityOptions = SecurityOptions.SSLImplicit;

        //    // Select folder
        //    client.SelectFolder("Inbox");
        //}




        public void LogOutMail()
        {
            this.imap.Close();
            MessageBox.Show("Email ile elaqe kesildi!");
        }





    }




       







}

