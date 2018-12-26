using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace luval.outlook.exporter
{
    public class OutlookExplorer
    {
        public OutlookExplorer()
        {
            
        }

        public void ReadAllMailItems()
        {
            var outlookApplication = new Application();
            var outlookNamespace = outlookApplication.GetNamespace("MAPI");
            var inboxFolder = outlookNamespace.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
            var mailItems = inboxFolder.Items;
            var isLoggedOn = false;
            try
            {
                //outlookNamespace.Logon("Profile", Missing.Value, false, true);
                outlookNamespace.Logon(Missing.Value, Missing.Value, false, true);
                isLoggedOn = true;
                Console.WriteLine("Accounts: {0}", outlookNamespace.Accounts.Count);

                foreach(Account acc in outlookNamespace.Accounts)
                {
                    Console.WriteLine("{0}", acc.DisplayName);
                }

                foreach (MailItem item in mailItems)
                {
                    var sb = new StringBuilder();
                    //sb.AppendLine("From: " + item.SenderEmailAddress);
                    sb.AppendLine("To: " + item.To);
                    sb.AppendLine("CC: " + item.CC);
                    sb.AppendLine("Subject: " + item.Subject);
                    sb.AppendLine("");
                    sb.AppendLine("==========================================================================");
                    Marshal.ReleaseComObject(item);
                }
            }
            catch(System.Exception ex)
            {
                Console.WriteLine(ex);
            }
            finally
            {
                if(isLoggedOn)
                    outlookNamespace.Logoff();
                ReleaseComObject(mailItems);
                ReleaseComObject(inboxFolder);
                ReleaseComObject(outlookNamespace);
                ReleaseComObject(outlookApplication);
            }
        }

        private static void ReleaseComObject(object obj)
        {
            if (obj != null)
            {
                Marshal.ReleaseComObject(obj);
                obj = null;
            }
        }
    }
}
