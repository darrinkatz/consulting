using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace GenerateEmailAudit
{
    class Program
    {
        static void Main(string[] args)
        {
            var currentDirectory = Directory.GetCurrentDirectory();
            string[] filePaths = Directory.GetFiles(currentDirectory);

            var mailItems = IngestFilePaths(filePaths);

            GenerateReport(new MasterReport(), mailItems);
            GenerateReport(new DailyTrafficReport(), mailItems);
            GenerateReport(new SendersReport(), mailItems);
            GenerateReport(new RecipientsReport(), mailItems);
            GenerateReport(new ConversationsReport(), mailItems);
            GenerateReport(new ReplyStatsReport(), mailItems);
            GenerateReport(new WordStatsReport(), mailItems);
            GenerateReport(new AttachmentsReport(), mailItems);
            GenerateReport(new BehaviourReport(), mailItems);

            Console.WriteLine("Done. Press any key to exit.");
            Console.ReadKey();
        }

        private static List<MailItem> IngestFilePaths(string[] filePaths)
        {
            var mailItems = new List<MailItem>();

            foreach (var filePath in filePaths)
            {
                try
                {
                    if (Path.GetExtension(filePath).ToLower().Equals(".csv"))
                    {
                        var allLines = File.ReadAllLines(filePath);
                        var mailItemCount = 0;

                        foreach (var line in allLines)
                        {
                            var mailItem = MailItem.Create(line);

                            if (mailItem != null && !mailItems.Contains<MailItem>(mailItem, new MailItemEqualityComparer()))
                            {
                                mailItems.Add(mailItem);
                                mailItemCount++;
                            }
                        }

                        Console.WriteLine(string.Format("Ingested {0} emails from {1}", mailItemCount, filePath));
                        Console.WriteLine();
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(string.Format("Error Ingesting {0}: {1}", filePath, ex.Message));
                    Console.WriteLine();
                    // swallow exception and continue loop
                }
            }

            return mailItems;
        }

        private static void GenerateReport(IReport report, List<MailItem> mailItems)
        {
            var reportDirectoryPath = string.Format(@"{0}\reports", Directory.GetCurrentDirectory());
            var reportFilePath = string.Format(@"{0}\{1}", reportDirectoryPath, report.GetReportFileName());
            
            try
            {
                if (!Directory.Exists(reportDirectoryPath))
                {
                    Directory.CreateDirectory(reportDirectoryPath);
                }

                using (StreamWriter file = new StreamWriter(reportFilePath, false))
                {
                    var reportData = report.Create(mailItems);

                    foreach (var line in reportData)
                    {
                        file.WriteLine(line);
                    }
                }

                Console.WriteLine(string.Format("Generated {0}", reportFilePath));
                Console.WriteLine();
            }
            catch (Exception ex)
            {
                Console.WriteLine(string.Format("Error Generating {0}: {1}", reportFilePath, ex.Message));
                Console.WriteLine();
                // swallow exception and continue
            }
        }
    }
}
