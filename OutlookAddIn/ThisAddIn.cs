using System;
using System.Net;
using System.IO;
using Newtonsoft.Json;
using System.Web;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace OutlookAddIn
{
    public partial class ThisAddIn
    {
        Outlook.Inspectors inspectors;
        //Creates inspector when Outlook opens
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            inspectors = this.Application.Inspectors;
            inspectors.NewInspector +=
                new Microsoft.Office.Interop.Outlook.InspectorsEvents_NewInspectorEventHandler(Inspectors_NewInspector);
        }

        //Gets data and adds to subject and body
        void Inspectors_NewInspector(Microsoft.Office.Interop.Outlook.Inspector Inspector)
        {
            Outlook.MailItem mailItem = Inspector.CurrentItem as Outlook.MailItem;
            WebRequest request = WebRequest.Create("http://brianeno.needsyourhelp.org/draw");
            WebResponse response = request.GetResponse();
            Stream dataStream = response.GetResponseStream();
            StreamReader reader = new StreamReader(dataStream);
            string responseFromServer = reader.ReadToEnd();

            dynamic data = JsonConvert.DeserializeObject(responseFromServer);
            string cardnumber = data.cardnumber;
            string strategy = data.strategy;

            string subject = "[OUTREACH][#" + cardnumber + "]";
            string body = "<body>Your strategy today is: " + strategy + "<br><br><body>";

            mailItem.Subject = subject + mailItem.Subject;
            mailItem.HTMLBody = body + mailItem.HTMLBody;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see http://go.microsoft.com/fwlink/?LinkId=506785
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
