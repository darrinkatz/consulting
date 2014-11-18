using System.Collections.Generic;

namespace GenerateEmailAudit
{
    public class MailItemEqualityComparer : IEqualityComparer<MailItem>
    {
        public bool Equals(MailItem x, MailItem y)
        {
            return x.ConversationIndex.Equals(y.ConversationIndex);
        }

        public int GetHashCode(MailItem obj)
        {
            return obj.GetHashCode();
        }
    }
}
