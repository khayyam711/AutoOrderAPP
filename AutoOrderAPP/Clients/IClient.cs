using System;
using System.Collections.Generic;
using System.Text;

namespace AutoOrderAPP.Clients
{
    public interface IClient
    {  
        public string MarketName { get; set; }
        public string MarketStickerNo { get; set; }
        public string MarketDepoNo { get; set; }
        
        public void CreateRootFolders(string date, string specialname, string routefolder, List<Client> market_names);

        public  List<Client> GetAllFolderName(string routefolder,string weekofday);
    }
}
