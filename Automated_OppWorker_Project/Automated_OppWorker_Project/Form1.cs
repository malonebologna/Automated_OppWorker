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
            OppEntry_Html_Creator.EntryManager entryManager = new EntryManager();
            OppEntry_Html_Creator.HTMLCreator htmlCreator = new HTMLCreator();
            Opp_Email_Manager.EmailCreator emailCreator = new EmailCreator();

            List<string[]> oppEntryList = entryManager.GetEntriesFromFile(@"C:\Users\FTCC\Desktop\BDSOpportunityAssignment.xlsx");
            foreach (string[] oppEntryDetails in oppEntryList)
            {
                string html = htmlCreator.UpdateEmailHTML(oppEntryDetails, @"C:\Users\FTCC\Desktop\OppEmailHTML.html");

                emailCreator.CreateEmailWithContacts(emailCreator.GetContactsFromFile(oppEntryDetails[0],@"C:\Users\FTCC\Desktop\OppTester")
                ,oppEntryDetails[3] , html, oppEntryDetails[2]);
            }

            

        
        }
    }
}
