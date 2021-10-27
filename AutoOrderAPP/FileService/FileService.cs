using System;
using System.Collections.Generic;
using System.Text;
using Limilabs.Client.IMAP;
using Limilabs.Mail;
using Limilabs.Mail.MIME;
using Imap = Limilabs.Client.IMAP.Imap;
using System.Windows;
using AutoOrderAPP.Clients;


namespace AutoOrderAPP.Orders


{ namespace AutoOrderAPP.FileService
{
    public class FileService :IFileService
    {

        public string OrderNo { get; set; }
        public string OrderDate { get; set; }
        public string BrendName { get; set; }
        public string OrderCari { get; set; }
        public string OrderDepoCode { get; set; }
        public string[] FN_split { get; set; }

        public bool DateConfirm = false;
        string temp = "";
        int count = 0;

        public int DownloadFromMail(string login,string key,string date,Client client, List<Client> rootfolders,string mailfoldername)
        {
                
                MessageBox.Show("Yuklenme baslayir!");
                if(mailfoldername!=null)
                {
                    using (Imap imap = new Imap())
                    {
                        imap.ConnectSSL("imap.yandex.ru" , 993);
                        imap.UseBestLogin(login, key);
                       
                        try
                        {
                            imap.Select(mailfoldername);                
                            List<long> uids = imap.Search(Flag.Unseen);

                            if (uids.Count!=0)
                            {
                                foreach (long uid in uids)
                                {
                                    IMail email = new MailBuilder().CreateFromEml(imap.GetMessageByUID(uid));

                                    if (email.ReturnPath == client.SenderMail)
                                    {
                                        foreach (MimeData mime in email.Attachments)
                                        {

                                            temp = this.ChangeFileName(mime.FileName, client.SplitChar, date);
                                            
                                            if (DateControl(date, this.OrderDate))
                                            {
                                                mime.Save(mailfoldername+"\\"+date + "\\" + mailfoldername + "\\" + MoveToFolder(OrderDepoCode, rootfolders) + "\\" + temp);
                                                count++;
                                                
                                            }
                                            else
                                            {
                                                imap.MarkMessageUnseenByUID(uid);
                                            }

                                        }
                                    }
                                    else
                                    {
                                        imap.MarkMessageUnseenByUID(uid);
                                    }
                                }
                            }
                            else
                            {
                                MessageBox.Show("Seçilmis Qovluqda yeni mesaj tapilmadi!");
                            }
                                                        

                        }
                        catch
                        {
                            MessageBox.Show("Sorguya uygun qovluq tapilmadi " +
                                             "\n Zəhmət olmasa daxil oldugunuz e-maildə " +
                                             "\n "+mailfoldername+" adlı qovluğun mövcudluğunu yoxlayın. ");
                           

                        }

                    }
                }
                return count;
        }

        public string ChangeFileName(string filneame, string splitchar, string date)
        {       
                try
                {
                    FN_split = filneame.Split(splitchar);
                    this.OrderCari = FN_split[3];
                    this.OrderDate = FN_split[1];
                    this.OrderDepoCode = FN_split[5].Split(new char[] { '.' })[0];
                    if (date == this.OrderDate)
                    {
                       
                        return DetectBrendName(this.OrderCari, filneame) + ".pdf";
                    }
                    else
                    {
                        OrderDepoCode = "-1";
                        return filneame;
                    }                     
                }    
                catch
                {
                    MessageBox.Show("Error Filename out off structur");
                    OrderDepoCode = "-1";
                    return filneame;
                }
        }

        public bool DateControl(string todaydate, string orderdate)
        {

            try
            {
                    if(todaydate == orderdate)
                    {
                        this.DateConfirm = true;
                        return true;
                    }
                    else
                    {
                        this.DateConfirm = false;
                        return false;
                    }
            }
            catch
            {
                MessageBox.Show("Please,check format of orderdate!It have to build dd.mm.yyyy format.You can find that file in " + Client.ErrorFolder + " folder.");
                return false;
            }
        }

        public string DetectBrendName(string cari, string filename)
        {
                switch (cari)
                {
                    case "2007":
                        this.BrendName = "HENKEL";
                        break;
                    case "1649":
                        this.BrendName = "GR";
                        break;
                    case "2147":
                        this.BrendName = "ARAQ";
                        break;
                    case "3901":
                        this.BrendName = "HILLSIDE";
                        break;
                    case "0145":
                        this.BrendName = "BELLA";
                        break;
                    case "0144":
                        this.BrendName = "HEINZ";
                        break;
                    case "1841":
                        this.BrendName = "WINE";
                        break;
                    case "1281":
                        this.BrendName = "SCJ";
                        break;
                    default:
                        {
                            this.BrendName = filename;
                            break;
                        }
                }

                return this.BrendName;
        }

        public string MoveToFolder(string orderdepocode, List<Client> rootfolders)
        {


                string path = Client.ErrorFolder;
                
                foreach (var folder in rootfolders)
                {   
                    if (orderdepocode == folder.MarketDepoNo)
                    {
                        path= (folder.MarketName + "_" + folder.MarketStickerNo);
                        break;
                    }
                    
                }
                return path;
            
        }

        public void UploadToServer(string dest_path)
        {
            throw new NotImplementedException();
        }

        public string ChangeFileName(string filename)
        {
            throw new NotImplementedException();
        }
    }
}
}