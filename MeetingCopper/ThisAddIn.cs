using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;


namespace MeetingCopper
{
    public partial class ThisAddIn
    {
         protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
          {
              return new Ribbon();
          }

        Outlook.Inspectors inspectors;
      
        


        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            inspectors = this.Application.Inspectors;
            inspectors.NewInspector +=
            new Microsoft.Office.Interop.Outlook.InspectorsEvents_NewInspectorEventHandler(Inspectors_NewInspector);
        }
        void Inspectors_NewInspector(Microsoft.Office.Interop.Outlook.Inspector Inspector)
        {
            Outlook.MailItem mailItem = Inspector.CurrentItem as Outlook.MailItem;
            Outlook.AppointmentItem meetingItem = Inspector.CurrentItem as Outlook.AppointmentItem;
            Type type = typeof(MeetingCopper.Ribbon);
            Ribbon ribbon = Globals.Ribbons.GetRibbon(type) as Ribbon;
            
            if (mailItem != null)
            {
                //ribbon.habilitaMail();
                ribbon.estado_meeting = false;
                ribbon.estado_mail = true;
            }

            if (meetingItem != null)
            {
                //MessageBox.Show("Meeting");
                //ribbon.habilitaMeeting();
                ribbon.estado_meeting = true;
                ribbon.estado_mail = false;
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Nota: Outlook ya no genera este evento. Si tiene código que 
            //    se debe ejecutar cuando Outlook se apaga, consulte https://go.microsoft.com/fwlink/?LinkId=506785
        }

        #region Código generado por VSTO

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido de este método con el editor de código.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
