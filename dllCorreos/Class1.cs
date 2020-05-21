using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace dllCorreos
{
    public class Class1
    {

        public Boolean SendEmailWithOutlook(string mailDirection, string mailSubject, string mailContent)
        {
            try
            {
                var oApp = new Outlook.Application();

                Outlook.NameSpace ns = oApp.GetNamespace("MAPI");
                var f = ns.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderInbox);

                System.Threading.Thread.Sleep(1000);

                var mailItem = (Outlook.MailItem)oApp.CreateItem(Outlook.OlItemType.olMailItem);
                mailItem.Subject = mailSubject;
                mailItem.HTMLBody = mailContent;
                mailItem.To = mailDirection;
                mailItem.Display(true);
                //mailItem.Send();

            }
            catch (Exception ex)
            {
                return false;
            }
            finally
            {
            }
            return true;
        }

    }
}
