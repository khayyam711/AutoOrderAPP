using AutoOrderAPP.Clients;
using System.Collections.Generic;
using Imap = Limilabs.Client.IMAP.Imap;

namespace AutoOrderAPP
{
    public interface IFileService
    {


        public int DownloadFromMail(string login, string key, string date,Client client, List<Client> rootfolders,string specialname);
        public bool DateControl(string todaydate,string orderdate );

        public string  MoveToFolder( string dest_path, List<Client> rootfolders);

        public  void  UploadToServer(string dest_path);
    }
}
