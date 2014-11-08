using System;

namespace GenerateEmailAudit
{
    public class MailItem
    {
        public string FolderPath { get; set; }
        public string[] Attachments { get; set; }
        public string BCC { get; set; }
        public string Body { get; set; }
        public string CC { get; set; }
        public string ConversationIndex { get; set; }
        public string ConversationTopic { get; set; }
        public DateTime CreationTime { get; set; }
        public DateTime LastModificationTime { get; set; }
        public string ReceivedByName { get; set; }
        public DateTime ReceivedTime { get; set; }
        public string[] Recipients { get; set; }
        public string SenderEmailAddress { get; set; }
        public string SenderName { get; set; }
        public DateTime SentOn { get; set; }
        public string Subject { get; set; }
        public string[] To { get; set; }
        public bool Unread { get; set; }

        private MailItem()
        {
        }

        public static MailItem Create(string line)
        {
            MailItem result = null;

            try
            {
                var columns = line.Split('\t');

                result = new MailItem()
                {
                    FolderPath = columns[0],
                    Attachments = columns[1].Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries),
                    BCC = columns[2],
                    Body = columns[3],
                    CC = columns[4],
                    ConversationIndex = columns[5],
                    ConversationTopic = columns[6],
                    CreationTime = DateTime.Parse(columns[7]),
                    LastModificationTime = DateTime.Parse(columns[8]),
                    ReceivedByName = columns[9],
                    ReceivedTime = DateTime.Parse(columns[10]),
                    Recipients = columns[11].Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries),
                    SenderEmailAddress = columns[12],
                    SenderName = columns[13],
                    SentOn = DateTime.Parse(columns[14]),
                    Subject = columns[15],
                    To = columns[16].Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries),
                    Unread = bool.Parse(columns[17]),
                };
            }
            catch (Exception ex)
            {
                Console.WriteLine(string.Format("Error Ingesting {0}: {1}", line, ex.Message));
                Console.WriteLine();
                // swallow exception and continue
            }

            return result;
        }
    }
}
