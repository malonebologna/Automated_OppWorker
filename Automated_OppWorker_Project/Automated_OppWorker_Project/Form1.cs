using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OppEntry_Html_Creator;
using Opp_Email_Manager;

namespace Automated_OppWorker_Project
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void goButton_Click(object sender, EventArgs e)
        {
            string oppContactsFolderLocation = @"C:\Users\FTCC\Desktop\Opportunities";
            string machformEntryFile = @"C:\Users\FTCC\Desktop\OppTester\BDSOpportunityAssignment.xlsx";
            string htmlTemplateFile = @"C:\Users\FTCC\Desktop\OppTester\OppEmailHTML.html";

            OppEntry_Html_Creator.EntryManager entryManager = new EntryManager();
            OppEntry_Html_Creator.HTMLCreator htmlCreator = new HTMLCreator();
            Opp_Email_Manager.EmailCreator emailCreator = new EmailCreator();

            List<string[]> oppList = entryManager.GetEntriesFromFile(machformEntryFile);
            foreach (string[] oppEntryFormData in oppList)
            {
                string solicationNumber = oppEntryFormData[0];
                string oppTitle = oppEntryFormData[3];
                string bdsEmailAddress = oppEntryFormData[2];
                string html = htmlCreator.UpdateEmailHTML(oppEntryFormData, htmlTemplateFile);

                emailCreator.CreateEmailWithContacts(
                emailCreator.GetContactsFromFile(solicationNumber, oppContactsFolderLocation)
                , oppTitle, html, bdsEmailAddress);
            }

            

        
        }
    }
}
