﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;
using System.Drawing;
using MeetingCopper.Properties;
using System.Xml.Linq;

// TODO:  Siga estos pasos para habilitar el elemento (XML) de la cinta de opciones:

// 1: Copie el siguiente bloque de código en la clase ThisAddin, ThisWorkbook o ThisDocument.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon();
//  }

// 2. Cree métodos de devolución de llamada en el área "Devolución de llamadas de la cinta de opciones" de esta clase para controlar acciones del usuario,
//    como hacer clic en un botón. Nota: si ha exportado esta cinta de opciones desde el diseñador de la cinta de opciones,
//    mueva el código de los controladores de eventos a los métodos de devolución de llamada y modifique el código para que funcione con el
//    modelo de programación de extensibilidad de la cinta de opciones (RibbonX).

// 3. Asigne atributos a las etiquetas de control del archivo XML de la cinta de opciones para identificar los métodos de devolución de llamada apropiados en el código.  

// Para obtener más información, vea la documentación XML de la cinta de opciones en la Ayuda de Visual Studio Tools para Office.


namespace MeetingCopper
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;


        public Ribbon()
        {

        }
        public Bitmap MinutaIcon(Microsoft.Office.Core.IRibbonControl control) => Resources.acuerdo00;
        public Bitmap RutinaIcon(Microsoft.Office.Core.IRibbonControl control) => Resources.rutina00;
        public Bitmap MeetingIcon(Microsoft.Office.Core.IRibbonControl control) => Resources.meeting2;


        public void OnClick(Office.IRibbonControl control)
        {
            
        }

        public void NuevaCita(Office.IRibbonControl control)
        {
            try
            {
                Microsoft.Office.Interop.Outlook.Application app = Globals.ThisAddIn.Application;
                Microsoft.Office.Interop.Outlook.AppointmentItem newCita = (Microsoft.Office.Interop.Outlook.AppointmentItem)
                app.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olAppointmentItem);
                if (newCita != null)
                {
                    RichTextBox rtb = new RichTextBox();
                    rtb.Rtf = System.Text.Encoding.UTF8.GetString(newCita.RTFBody);
                    rtb.Select(rtb.TextLength, 0);

                    rtb.LoadFile("C:\\Users\\Administrator\\Downloads\\MC_NuevaReunion.rtf");

                    newCita.RTFBody = System.Text.Encoding.UTF8.GetBytes(rtb.Rtf);

                    newCita.MeetingStatus = Microsoft.Office.Interop.Outlook.OlMeetingStatus.olMeeting;
                    
                    newCita.Start = DateTime.Now.AddHours(2);
                    newCita.End = DateTime.Now.AddHours(3);
                    newCita.Location = "Elija la ubicación de la Reunión";
                    newCita.Subject = "Reunión Template";
                    newCita.Recipients.Add("Seleccione los Destinatarios");
                    newCita.Display(true);
                    newCita.AllDayEvent = false;
                }                
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops, ha ocurrido el siguiente error:  " + ex.Message);
            }
        }

        public void NuevaRutina(Office.IRibbonControl control)
        {
            try
            {
                Microsoft.Office.Interop.Outlook.Application app = Globals.ThisAddIn.Application;
                Microsoft.Office.Interop.Outlook.AppointmentItem newCita = (Microsoft.Office.Interop.Outlook.AppointmentItem)
                app.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olAppointmentItem);
                if (newCita != null)
                {
                    newCita.MeetingStatus = Microsoft.Office.Interop.Outlook.OlMeetingStatus.olMeeting;

                    RichTextBox rtb = new RichTextBox();
                    rtb.Rtf = System.Text.Encoding.UTF8.GetString(newCita.RTFBody);
                    rtb.Select(rtb.TextLength, 0);

                    rtb.LoadFile("C:\\Users\\Administrator\\Downloads\\MC_NuevaRutina.rtf");

                    newCita.RTFBody = System.Text.Encoding.UTF8.GetBytes(rtb.Rtf);

                    newCita.Start = DateTime.Now.AddHours(2);
                    newCita.End = DateTime.Now.AddHours(3);
                    newCita.Location = "Elija la ubicación de la Reunión";
                    newCita.Subject = "Reunión Template";
                    newCita.Recipients.Add("Seleccione los Destinatarios");

                    newCita.Display(true);
                    newCita.AllDayEvent = false;

                    
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops, ha ocurrido el siguiente error:  " + ex.Message);
            }
        }

        public void NuevaMinuta(Office.IRibbonControl control)
        {
            try
            {
                Microsoft.Office.Interop.Outlook.Application app = Globals.ThisAddIn.Application;
                Microsoft.Office.Interop.Outlook.MailItem newMail = (Microsoft.Office.Interop.Outlook.MailItem)
                app.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);
                if (newMail != null)
                {
                    string HTMLTemplate = File.ReadAllText("C:\\Users\\Administrator\\Downloads\\MC_NuevaMinuta.html");

                    newMail.Subject = "Template Minutas de Acciones";
                    newMail.HTMLBody = HTMLTemplate;
                    newMail.To = "Seleccione sus Destinatarios";
                    Microsoft.Office.Interop.Outlook.Recipients sentTo = newMail.Recipients;
                    sentTo.ResolveAll();
                    newMail.Display(true);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops, ha ocurrido el siguiente error:  " + ex.Message);
            }
        }

        #region Miembros de IRibbonExtensibility

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("MeetingCopper.Ribbon.xml");
        }

        #endregion

        #region Devoluciones de llamada de la cinta de opciones
        //Cree métodos de devolución de llamada aquí. Para obtener más información sobre la adición de métodos de devolución de llamada, visite https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        #endregion

        #region Asistentes

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
