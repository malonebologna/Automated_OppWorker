using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace OppEntry_Html_Creator
{
    public class HTMLCreator
    {
        public string UpdateEmailHTML(string[] oppDetails, string HtmlTemplateFileLocation)
        {
            StreamReader reader = new StreamReader(HtmlTemplateFileLocation);
            string oppEmailHtml = reader.ReadToEnd();

            oppEmailHtml = oppEmailHtml.Replace("##oppTitle##", oppDetails[3]);
            oppEmailHtml = oppEmailHtml.Replace("##solNum##", oppDetails[0]);
            oppEmailHtml = oppEmailHtml.Replace("##agency##", oppDetails[4]);
            oppEmailHtml = oppEmailHtml.Replace("##office##", oppDetails[5]);
            oppEmailHtml = oppEmailHtml.Replace("##location##", oppDetails[6]);
            oppEmailHtml = oppEmailHtml.Replace("##noticeType##", oppDetails[7]);
            oppEmailHtml = oppEmailHtml.Replace("##responseDate##", oppDetails[10]);
            oppEmailHtml = oppEmailHtml.Replace("##setAside##", oppDetails[11]);
            oppEmailHtml = oppEmailHtml.Replace("##naics##", oppDetails[12]);
            oppEmailHtml = oppEmailHtml.Replace("##linkToOpp##", oppDetails[14]);
            oppEmailHtml = oppEmailHtml.Replace("##synopsis##", oppDetails[15]);
            oppEmailHtml = oppEmailHtml.Replace("##contractingOfficeAddress##", oppDetails[16]);
            oppEmailHtml = oppEmailHtml.Replace("##pop##", oppDetails[17]);
            oppEmailHtml = oppEmailHtml.Replace("##primaryPOC##", oppDetails[18]);
            oppEmailHtml = oppEmailHtml.Replace("##bds##", oppDetails[1]);
            oppEmailHtml = oppEmailHtml.Replace("##bdsEmail##", oppDetails[2]);

            return oppEmailHtml;

        }
    }
}
