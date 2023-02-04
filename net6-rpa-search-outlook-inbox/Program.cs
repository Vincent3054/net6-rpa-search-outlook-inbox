using Microsoft.Office.Interop.Outlook;
using System.Diagnostics;
using Outlook = Microsoft.Office.Interop.Outlook;

Action<string> printTitle = (s) =>
{
    Console.ForegroundColor = ConsoleColor.Yellow;
    Console.WriteLine();
    Console.WriteLine(s);
    Console.ResetColor();
};

printTitle("人資系統申請通知'");

SearchInbox(senderName: "EIP@chailease.com.tw");

printTitle("請假申請流程");
SearchInbox(subjectKeywd: "請假申請流程");

printTitle("加班核銷流程");
SearchInbox(subjectKeywd: "加班核銷流程");

printTitle("一個月內的請假申請流程");
SearchInbox("EIP@chailease.com.tw", "請假申請流程", DateTime.Now.AddMonths(-1), DateTime.Now);

printTitle("一個月內的加班核銷流程");
SearchInbox("EIP@chailease.com.tw", "加班核銷流程", DateTime.Now.AddMonths(-1), DateTime.Now);

void SearchInbox(string senderName = null, string subjectKeywd = null, DateTime? startTime = null, DateTime? endTime = null)
{
    if (Process.GetProcessesByName("OUTLOOK").Length > 0)
    {
        var app = new Outlook.Application();
        var ns  = app.GetNamespace("MAPI");
        var inbox = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox); //最外層收件夾
        var todoinbox = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox); //待辦事項資料夾。
        var items = inbox.Items;
        var conds = new List<string>();
        Func<string, string> escape = (s) => s.Replace("'", "''");
        if (!string.IsNullOrEmpty(senderName))
        {
            conds.Add(@$"(""urn:schemas:httpmail:sendername"" = '{escape(senderName)}')");
        }
        if (!string.IsNullOrEmpty(subjectKeywd))
        {
            conds.Add(@$"(""urn:schemas:httpmail:subject"" LIKE '%{escape(subjectKeywd)}%')");
        }
        if (startTime != null)
        {
            conds.Add(@$"(""urn:schemas:httpmail:datereceived"" > '{startTime.Value:yyyy-MM-dd HH:mm:ss}')");
        }
        if (endTime != null)
        {
            conds.Add(@$"(""urn:schemas:httpmail:datereceived"" < '{endTime.Value:yyyy-MM-dd HH:mm:ss}')");
        }
        var filterString = "@SQL=" + string.Join(" AND ", conds.ToArray());
        var filterd = items.Restrict(filterString);
        if (filterd.Count == 0)
        {
            Console.WriteLine("Not Found");
        }
        else
        {
            foreach (var item in items.Restrict(filterString))
            {
                var mailItem = item as Outlook.MailItem;
                if (mailItem != null)
                {
                    Console.WriteLine($"{mailItem.SentOn:yyyy-MM-dd HH:mm:ss} / {mailItem.SenderName} / {mailItem.Subject}");
                }
            }
        }
    }
    else
    {
        Console.WriteLine("Outlook is not running");
    }
}