using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;

namespace OutlookLocationImage
{
    public partial class ThisAddIn
    {
        private Inspectors inspectors;
        private MeetingItem meetingItem;
        private AppointmentItem appointmentItem;
        private string CurrentLocation;

        private const string Location = "Location";
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            inspectors = Globals.ThisAddIn.Application.Inspectors;
            inspectors.NewInspector += Inspectors_NewInspector;
        }

        private void Inspectors_NewInspector(Inspector Inspector)
        {
            if (Inspector.CurrentItem is AppointmentItem)
            {
                appointmentItem = Inspector.CurrentItem;
                CurrentLocation = appointmentItem.Location;
                appointmentItem.PropertyChange += AppointmentItem_PropertyChange;
            }
            //if (Inspector.CurrentItem is MeetingItem)
            //{
            //    meetingItem = Inspector.CurrentItem;
            //    meetingItem.PropertyChange += MeetingItem_PropertyChange;
            //}
            //else if (Inspector.CurrentItem is AppointmentItem)
            //{
            //    appointmentItem = Inspector.CurrentItem;
            //    appointmentItem.PropertyChange += AppointmentItem_PropertyChange;
            //}
        }
        private void MeetingItem_PropertyChange(string Name)
        {
            if (Name == Location)
            {
                MessageBox.Show(meetingItem.PropertyAccessor.GetProperty(Name));
            }
        }
        private void AppointmentItem_PropertyChange(string Name)
        {
            if (Name == Location && appointmentItem.Location != CurrentLocation)
            {
                MessageBox.Show(appointmentItem.Location);
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
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
