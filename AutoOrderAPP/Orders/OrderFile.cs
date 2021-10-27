using System;
using System.Collections.Generic;
using System.Text;
using Limilabs.Client.IMAP;
using Limilabs.Mail;
using Limilabs.Mail.MIME;
using Imap = Limilabs.Client.IMAP.Imap;
using System.Windows;

namespace AutoOrderAPP.Orders
{
    public class OrderFile
    {

        public string OrderNo { get; set; }
        public string OrderDate { get; set; }
        public string BrendName { get; set; }
        public string OrderCari { get; set; }
        public string OrderDepoCode { get; set; }
        public string[] FN_split { get; set; }
     

       
    }
}
