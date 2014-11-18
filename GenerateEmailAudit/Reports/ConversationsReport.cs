using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GenerateEmailAudit
{
    class ConversationsReport : IReport
    {
        public string GetReportFileName()
        {
            return "ConversationsReport.txt";
        }

        public string[] Create(List<MailItem> mailItems)
        {
            var result = new List<string>();
            var sb = new StringBuilder();

            sb.Append(string.Format("{0}\t", "Thread"));
            sb.Append(string.Format("{0}\t", "Total Replies"));
            sb.Append(string.Format("{0}\t", "Replies from User"));
            result.Add(sb.ToString());
            sb.Clear();

            result.Add(string.Empty);

            var conversations = GetConversations(mailItems);

            foreach (var conversation in conversations)
            {
                sb.Append(string.Format("{0}\t", conversation.ConversationTopic));
                sb.Append(string.Format("{0}\t", conversation.TotalReplies));
                sb.Append(string.Format("{0}\t", conversation.RepliesFromUser));
                result.Add(sb.ToString());
                sb.Clear();
            }

            return result.ToArray();
        }

        private static List<Conversation> GetConversations(List<MailItem> mailItems)
        {
            var result = new List<Conversation>();

            foreach (var mailItem in mailItems)
            {
                var conversation = result.FirstOrDefault(t => t.ConversationTopic.Equals(mailItem.ConversationTopic));

                if (conversation != null)
                {
                    conversation.TotalReplies++;
                    conversation.RepliesFromUser += mailItem.CurrentUserAddressEntry.Equals(mailItem.SenderName) ? 1 : 0;
                }
                else
                {
                    result.Add(new Conversation()
                    {
                        ConversationTopic = mailItem.ConversationTopic,
                        TotalReplies = 0,
                        RepliesFromUser = 0,
                    });
                }
            }

            return result;
        }
    }
}
