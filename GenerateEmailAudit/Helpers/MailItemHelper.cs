using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GenerateEmailAudit
{
    public class MailItemHelper
    {
        public static List<MailItem> GetAllEmailsSent(List<MailItem> mailItems)
        {
            return mailItems.Where(mi =>
                mi.SenderName.Equals(mi.CurrentUserAddressEntry)
                ).ToList();
        }

        public static List<MailItem> GetAllEmailsReceived(List<MailItem> mailItems)
        {
            return mailItems.Where(mi =>
                mi.ReceivedByName.Equals(mi.CurrentUserAddressEntry)
                ).ToList();
        }
    }
}
