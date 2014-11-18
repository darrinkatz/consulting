using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GenerateEmailAudit
{
    public class Conversation
    {
        public string ConversationTopic { get; set; }
        public int TotalReplies { get; set; }
        public int RepliesFromUser { get; set; }
    }
}
