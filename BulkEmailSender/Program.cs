using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BulkEmailSender
{
    class Program
    {
        static void Main(string[] args)
        {
            BulkEmailSender bulkEmailSender = new BulkEmailSender();
            bulkEmailSender.SendBulkEmails();    
        }
    }
}
