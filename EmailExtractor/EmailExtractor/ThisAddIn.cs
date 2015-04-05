using Microsoft.Office.Interop.Outlook;
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;

namespace EmailExtractor
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Application.MAPILogonComplete += new ApplicationEvents_11_MAPILogonCompleteEventHandler(MAPILogonCompleteEventHandler);
        }

        void MAPILogonCompleteEventHandler()
        {
            var username = this.Application.Session.CurrentUser.Name;

            // TODO: make parameter
            var directory = @"C:\ProgramData\Inbox_Analysis_Tool\";
            Directory.CreateDirectory(directory);

            var pathFormat = string.Format("{0}{1}_{2}_{3}",
                directory,
                new Regex(@"[^A-Za-z]+").Replace(username, "_").ToLower(),
                DateTime.Now.ToString("yyyyMMddHHmmss"),
                "{0}"
                );

            try
            {
                EmailExtractor.ProcessFolder(pathFormat, username, Application.Session.DefaultStore.GetRootFolder() as Folder);
            }
            catch (System.Exception ex)
            {
                using (var filestream = File.AppendText(string.Format(pathFormat, "errors.txt")))
                {
                    filestream.WriteLine(DateTime.Now.ToString());
                    filestream.WriteLine(ex);
                }
            }
            finally
            {
                MailItem mailItem = Application.CreateItem(OlItemType.olMailItem);
                mailItem.Recipients.Add("darrinkatz@gmail.com");   // TODO: make parameter
                mailItem.Recipients.Add("darrin.katz@shawmedia.ca");   // TODO: make parameter
                mailItem.Subject = string.Format("Inbox_Analysis_Tool Files for {0}", username);
                foreach (var file in Directory.GetFiles(directory))
                {
                    mailItem.Attachments.Add(file);
                }
                ((Microsoft.Office.Interop.Outlook._MailItem)mailItem).Send();
                foreach(var file in Directory.GetFiles(directory))
                {
                    File.Delete(file);
                }
            }
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
