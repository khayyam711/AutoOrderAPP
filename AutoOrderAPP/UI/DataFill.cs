using AutoOrderAPP.Clients;
using System.Collections.Generic;

namespace AutoOrderAPP.UI
{
    partial class DataFill : MainWindow
    {
       

        public  void FormFillData(List<Client> clients_info)
        {

            foreach (var name  in clients_info)
            {
                folder_name.Items.Add(name.MarketName);
            }

        }



    }
}
