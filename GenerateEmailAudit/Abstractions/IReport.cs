using System.Collections.Generic;

namespace GenerateEmailAudit
{
    interface IReport
    {
        string GetReportFileName();
        string[] Create(List<MailItem> mailItems);
    }
}
