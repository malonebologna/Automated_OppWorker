using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace OppEntry_Html_Creator
{
    public class EntryManager
    {
        public List<string[]> GetEntriesFromFile(string filePath)
        {            
            Excel.Application oXL;
            Excel._Worksheet oSheet;
            Excel.Range range;
            oXL = new Excel.Application();
            oXL.Visible = false;
            Excel.Workbook oWB = oXL.Workbooks.Open(filePath,
                    0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "",
                    true, false, 0, true, false, false);
            oSheet = (Excel._Worksheet)oWB.ActiveSheet;
            range = oSheet.UsedRange;
            int rowCount = range.Rows.Count;
            int currentRow = 2;

            
            List<string[]> oppList = new List<string[]>();

        while (currentRow <= rowCount)
        {
            string[] currentOppDetails = new string[19];
            currentOppDetails[0] = range.Cells[currentRow, 9].Text; 
            currentOppDetails[1] = range.Cells[currentRow, 6].Text;
            currentOppDetails[2] = range.Cells[currentRow, 7].Text;
            currentOppDetails[3] = range.Cells[currentRow, 8].Text;
            currentOppDetails[4] = range.Cells[currentRow, 10].Text;
            currentOppDetails[5] = range.Cells[currentRow, 11].Text;
            currentOppDetails[6] = range.Cells[currentRow, 12].Text;
            currentOppDetails[7] = range.Cells[currentRow, 13].Text;
            currentOppDetails[8] = range.Cells[currentRow, 14].Text;
            currentOppDetails[9] = range.Cells[currentRow, 15].Text;
            currentOppDetails[10] = range.Cells[currentRow, 16].Text;
            currentOppDetails[11] = range.Cells[currentRow, 17].Text;
            currentOppDetails[12] = range.Cells[currentRow, 18].Text;
            currentOppDetails[13] = range.Cells[currentRow, 20].Text;
            currentOppDetails[14] = range.Cells[currentRow, 21].Text;
            currentOppDetails[15] = range.Cells[currentRow, 22].Text;
            currentOppDetails[16] = range.Cells[currentRow, 23].Text;
            currentOppDetails[17] = range.Cells[currentRow, 24].Text;
            currentOppDetails[18] = range.Cells[currentRow, 25].Text;


            oppList.Add(currentOppDetails);
            currentRow++;
        }

        oXL.Quit();
        return oppList;
            
        }

        public void PrintEntriesToConsole(List<string[]> oppList)
        {
            foreach (string[] oppDetails in oppList)
            {
                foreach (string detail in oppDetails)
                {
                    Console.WriteLine(detail);
                }
            }
        }
    }
}
