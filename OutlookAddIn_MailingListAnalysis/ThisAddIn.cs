using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace OutlookAddIn_MailingListAnalysis
{
    public partial class ThisAddIn
    {
        Outlook.Inspectors insp;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            insp = this.Application.Inspectors;
            insp.NewInspector += new Microsoft.Office.Interop.Outlook.InspectorsEvents_NewInspectorEventHandler(Inspectors_NewInspector);
        }

        void Inspectors_NewInspector(Microsoft.Office.Interop.Outlook.Inspector insp)
        {
            Outlook.MailItem mailItem = insp.CurrentItem as Outlook.MailItem;
            if(mailItem !=null)
            {
                if (mailItem.EntryID == null)
                {
                    mailItem.Subject = "I updated this text from code.";
                    mailItem.Body = "This code works!";

                }

            }


        }
        

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
