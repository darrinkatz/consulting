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
        private static string _delimiter = "\t";

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            var username = this.Application.Session.CurrentUser.Name;

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
                var rootFolder = Application.Session.DefaultStore.GetRootFolder() as Folder;
                this.ProcessFolder(pathFormat, username, rootFolder);
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
                // TODO: email everything in the folder
            }
        }

        private void ProcessFolder(string pathFormat, string username, Folder folder)
        {
            var table = folder.GetTable();
            table.Columns.Add("http://schemas.microsoft.com/mapi/proptag/0x10F4000B");

            while (!table.EndOfTable)
            {
                var nextRow = table.GetNextRow();

                if (nextRow.GetValues()[4] == "IPM.Note")
                {
                    var sb = new StringBuilder();

                    sb.Append(username);
                    sb.Append(_delimiter);

                    sb.Append(folder.FolderPath);
                    sb.Append(_delimiter);

                    if (nextRow["http://schemas.microsoft.com/mapi/proptag/0x10F4000B"] != null)
                    {
                        sb.Append(nextRow["http://schemas.microsoft.com/mapi/proptag/0x10F4000B"].ToString());
                        sb.Append(_delimiter);
                    }

                    using (var outfile = File.AppendText(string.Format(pathFormat, "data.csv")))
                    {
                        outfile.WriteLine(sb.ToString());
                    }
                }
            }

            for (int i = 1; i < folder.Folders.Count; i++)
            {
                var childFolder = folder.Folders[i] as Folder;

                this.ProcessFolder(pathFormat, username, childFolder);
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
