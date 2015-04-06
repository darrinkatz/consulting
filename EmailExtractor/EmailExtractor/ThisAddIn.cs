using Microsoft.Office.Interop.Outlook;
using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace EmailExtractor
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.ItemLoad += ThisApplication_ItemLoad;
        }

        void ThisApplication_ItemLoad(object Item)
        {
            // TODO: make parameter for reporting period
            var reportStartDate = new DateTime(2015, 4, 5);
            var reportEndDate = new DateTime(2015, 4, 17).AddDays(1);

            if (reportStartDate < DateTime.Now && DateTime.Now < reportEndDate)
            {
                var username = this.Application.Session.CurrentUser.Name;

                // TODO: make parameter
                var directory = @"C:\ProgramData\Inbox_Analysis_Tool\";
                Directory.CreateDirectory(directory);

                var pathFormat = string.Format("{0}{1}_{2}",
                    directory,
                    new Regex(@"[^A-Za-z]+").Replace(username, "_").ToLower(),
                    "{0}"
                    );

                try
                {
                    var timeSinceLastSnooze = TimeSpan.MaxValue;
                    foreach (var file in Directory.GetFiles(directory))
                    {
                        if (file.Contains("snooze"))
                        {
                            var snoozeTimeSpan = DateTime.Now - File.GetCreationTime(file);

                            if (snoozeTimeSpan < timeSinceLastSnooze)
                            {
                                timeSinceLastSnooze = snoozeTimeSpan;
                            }
                        }
                    }

                    if (timeSinceLastSnooze > new TimeSpan(1, 0, 0))
                    {
                        DateTime effectiveStartDate;
                        var timestampFilename = string.Format(pathFormat, "timestamp.txt");
                        
                        if (File.Exists(timestampFilename))
                        {
                            effectiveStartDate = File.GetCreationTime(timestampFilename);
                        }
                        else
                        {
                            effectiveStartDate = reportStartDate;
                        }

                        if (DateTime.Now - effectiveStartDate > new TimeSpan(24, 0, 0))
                        {
                            var caption = "Inbox Analysis Tool";
                            var message = "The Inbox Analysis Tool needs to create a report every 24 hours, and it will only take a few minutes. Click 'Yes' if that's okay or 'No' to snooze for 1 hour.";
                            
                            var result = MessageBox.Show(message, caption, MessageBoxButtons.YesNo);

                            if (result == DialogResult.Yes)
                            {
                                var rootFolder = Application.Session.DefaultStore.GetRootFolder() as Folder;
                                var filter = string.Format("urn:schemas:httpmail:datereceived >= '{0} 12:00AM' And urn:schemas:httpmail:datereceived <= '{1} 12:00AM'",
                                    effectiveStartDate.ToString("MM.dd.yyyy"),
                                    reportEndDate.ToString("MM.dd.yyyy")
                                    );
                                var scope = string.Format("'{0}'", rootFolder.FolderPath.Replace("\\\\", "\\"));
                                var table = Application.AdvancedSearch(scope, filter, true, "tag").GetTable();
                                EmailExtractor.ProcessTable(Application.GetNamespace("MAPI"), table, string.Format(pathFormat, string.Format("data_{0}.csv", DateTime.Now.ToString("yyyyMMddHHmmss"))), username);

                                this.EmailProgramData(username, directory);

                                if (!File.Exists(timestampFilename))
                                {
                                    File.CreateText(timestampFilename);
                                }

                                MessageBox.Show("All done! The Inbox Analysis Tool will run again in 24 hours.", caption);
                            }
                            else
                            {
                                File.CreateText(string.Format(pathFormat, string.Format("snooze_{0}.txt", DateTime.Now.ToString("yyyyMMddHHmmss"))));
                            }
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    using (var filestream = File.AppendText(string.Format(pathFormat, "errors.txt")))
                    {
                        filestream.WriteLine(DateTime.Now.ToString());
                        filestream.WriteLine(ex);
                    }

                    // TODO: figure out a way to email errors on a daily basis instead of every time
                    //this.EmailProgramData(username, directory);
                }
            }
        }

        private void EmailProgramData(string username, string directory)
        {
            MailItem mailItem = Application.CreateItem(OlItemType.olMailItem);
            mailItem.Recipients.Add("darrinkatz@gmail.com");   // TODO: make parameter
            mailItem.Recipients.Add("darrin.katz@shawmedia.ca");   // TODO: make parameter
            mailItem.Subject = string.Format("Inbox_Analysis_Tool Files for {0}", username);
            foreach (var file in Directory.GetFiles(directory))
            {
                if (file.Contains(".csv") || file.Contains("error"))
                {
                    mailItem.Attachments.Add(file);
                }
            }
            ((Microsoft.Office.Interop.Outlook._MailItem)mailItem).Send();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
