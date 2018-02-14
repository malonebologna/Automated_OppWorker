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
        public string UpdateEmailHTML(string[] oppEntryFormData, string HtmlTemplateFileLocation)
        {
            StreamReader reader = new StreamReader(HtmlTemplateFileLocation);
            string oppEmailHtml = reader.ReadToEnd();

            oppEmailHtml = oppEmailHtml.Replace("##oppTitle##", oppEntryFormData[3]);
            oppEmailHtml = oppEmailHtml.Replace("##solNum##", oppEntryFormData[0]);
            oppEmailHtml = oppEmailHtml.Replace("##agency##", oppEntryFormData[4]);
            oppEmailHtml = oppEmailHtml.Replace("##office##", oppEntryFormData[5]);
            oppEmailHtml = oppEmailHtml.Replace("##location##", oppEntryFormData[6]);
            oppEmailHtml = oppEmailHtml.Replace("##noticeType##", oppEntryFormData[7]);
            oppEmailHtml = oppEmailHtml.Replace("##responseDate##", oppEntryFormData[10]);
            oppEmailHtml = oppEmailHtml.Replace("##setAside##", oppEntryFormData[11]);
            oppEmailHtml = oppEmailHtml.Replace("##naics##", oppEntryFormData[12]);
            oppEmailHtml = oppEmailHtml.Replace("##linkToOpp##", oppEntryFormData[14]);
            oppEmailHtml = oppEmailHtml.Replace("##synopsis##", oppEntryFormData[15]);
            oppEmailHtml = oppEmailHtml.Replace("##contractingOfficeAddress##", oppEntryFormData[16]);
            oppEmailHtml = oppEmailHtml.Replace("##pop##", oppEntryFormData[17]);
            oppEmailHtml = oppEmailHtml.Replace("##primaryPOC##", oppEntryFormData[18]);
            oppEmailHtml = oppEmailHtml.Replace("##bds##", oppEntryFormData[1]);
            oppEmailHtml = oppEmailHtml.Replace("##bdsEmail##", oppEntryFormData[2]);

            return oppEmailHtml;

        }
    }
}
