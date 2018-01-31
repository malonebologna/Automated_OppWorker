using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Opp_Email_Manager
{
    public class EmailCreator
    {
        public string GetContactsFromFile(string solNum, string folderLocation)
        {
            Excel.Application oXL;
            Excel._Worksheet oSheet;
            Excel.Range range;
            oXL = new Excel.Application
            {
                Visible = false
            };
            Excel.Workbook oWB = oXL.Workbooks.Open(folderLocation + @"\" + solNum + ".xls",
                    0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "",
                    true, false, 0, true, false, false);
            oSheet = (Excel._Worksheet)oWB.ActiveSheet;
            range = oSheet.UsedRange;
            int rowCount = range.Rows.Count;
            int currentRow = 2;

            string contacts = "";
            while (currentRow <= rowCount)
            {
                contacts += range.Cells[currentRow, 6].Text + ", ";
                currentRow++;
            }
            
            oXL.Quit();
            return contacts;
        }

        public void CreateEmailWithContacts(string contacts, string title, string html, string bdsEmail)
        {
            Outlook.Application outlookApp = new Outlook.Application();
            Outlook.MailItem mailItem = (Outlook.MailItem)
                outlookApp.Application.CreateItem(Outlook.OlItemType.olMailItem);
            mailItem.Subject = "Opportunity: " + title;
            mailItem.BCC = contacts;
            mailItem.CC = bdsEmail;
            mailItem.Recipients.ResolveAll();
            mailItem.HTMLBody = html;
            mailItem.Display(true);

        }
    }
}
