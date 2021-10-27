using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text;
using System.Windows;

namespace AutoOrderAPP.Clients
{
    public class Client : IClient
    {
        //Properties
        public static string ErrorFolder = "UnRenamedFiles";
        public string MarketName { get; set; }
        public string MarketStickerNo { get; set; }
        public string MarketDepoNo { get; set; }
        public string SplitChar { get; set; }
        public string SenderMail { get; set; }
        public static string SpecialName { get; set; }
        public string FolderName { get; set; }


        //Contructors
        public Client()
        {

        }
        public Client(string routefolder) {
            SplitChar = routefolder + "split_char.txt";
            SenderMail = routefolder + "sendermail.txt";
            if (!Directory.Exists(SplitChar))
            {   
                SplitChar = File.ReadAllText(SplitChar).Trim();
            }
            if (!Directory.Exists(SenderMail))
            {
                SenderMail = File.ReadAllText(SenderMail).Trim();
            }
            

        }// constructor (arg)
        public Client(string marketname,string marketstickerno)
        {
            MarketName = marketname;
            MarketStickerNo = marketstickerno;
            FolderName = marketname + "_" + marketstickerno;

        } //Constructor(arg1,arg2)
        public Client(string marketname,string marketstickerno,string marketdepono)
        {
            MarketName = marketname;
            MarketStickerNo = marketstickerno;
            MarketDepoNo = marketdepono;
            FolderName = marketname + "_" + marketstickerno;


        } //Constructor(arg1, arg2, arg3 )

        //Methods
        public List<Client> GetAllFolderName(string routefolder, string weekOfDay)
        {
            List<Client> TodayFolders = new List<Client>();

            Client client = new Client(routefolder);
            string split_symb = client.SplitChar;
            string txtfile = routefolder + weekOfDay + ".txt";
            string[] allmarkets, temp = new string[] { };

            if (File.Exists(txtfile))
            {
                allmarkets = File.ReadAllLines(txtfile);
                for (int i = 0; i < allmarkets.Length; i++)
                {
                    
                    if (allmarkets[i] != "")
                    {   

                        temp = allmarkets[i].Split(split_symb) ;
                        
                        try
                        {
                            client = new Client(temp[0], temp[1], temp[2]);
                        }
                        catch
                        {
                            try
                            {
                                client = new Client(temp[0], temp[1]);
                            }
                            catch
                            {
                                MessageBox.Show("Please check structur of .txt file in Route Folder" +
                                                "\n It have to create (MARKETNAME_STICKERCODE " +
                                                "\n OR marketname_stickercode_depocode " +
                                                "\n \n '_' This sign  is important for \n " +
                                                "split marketname,stkicercode and depocoe.");
                            }

                        }

                        TodayFolders.Add(client);
                    }
                }
            }

            return TodayFolders;
        }

        public void CreateRootFolders(string date,string specialname,string routefolder,List<Client> market_names)
        {

            string specialpath = specialname+"\\" +date+ "\\"+specialname+"\\";
            string lastpath = "";

            
            if (!Directory.Exists(specialpath))
                {
                    Directory.CreateDirectory(specialpath);
                    foreach (var market in market_names)
                    {
                        lastpath = specialpath + market.FolderName+"\\Hazir";
                        if (!Directory.Exists(lastpath))
                        {
                             Directory.CreateDirectory(lastpath);
                             File.CreateText(specialpath+"\\"+market.MarketName+"OrderNo.txt");
                        }                        
                    }
                    Directory.CreateDirectory(specialpath+ErrorFolder);
                }    
                
            }

        
    }
}
