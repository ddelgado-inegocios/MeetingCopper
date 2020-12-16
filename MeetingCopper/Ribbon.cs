// <copyright solution="MeetingTools" project="MeetingTools" file="Ribbon.cs">
// 
//     TIAXA CHILE CONFIDENCIAL
//     --------------------------------------------------------------------------------------------
// 
//     NOTICE:  All information contained herein is, and remains the property of Tiaxa Chile,
//     The intellectual and technical concepts contained herein are proprietary to Tiaxa and
//     and are protected by trade secret or copyright law.
//     Dissemination of this information or reproduction of this material is strictly forbidden
//     unless prior written permission is obtained from Tiaxa.
// 
//     File Created  by Cristian Ronny Meneses González at 10-12-2020 17:42
//     File Modified by Cristian Ronny Meneses González at 16-12-2020 2:53
// 
//     Copyright ©2020 Tiaxa Chile. All rights reserved.
// 
// </copyright>
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using MeetingCopper.Properties;
using Microsoft.Office.Interop.Outlook;
using Application = Microsoft.Office.Interop.Outlook.Application;
using Office = Microsoft.Office.Core;

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

namespace MeetingCopper {

    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility {

        private Office.IRibbonUI _ribbon;

        #region Miembros de IRibbonExtensibility

        public string GetCustomUI(string ribbonId) {
            //MessageBox.Show(ribbonId);
            //return GetResourceText("MeetingCopper.Ribbon.xml");

            switch (ribbonId) {
                case "Microsoft.Outlook.Appointment":
                    return GetResourceText("MeetingCopper.RibbonMeeting.xml");
                case "Microsoft.Outlook.Mail.Compose":
                    return GetResourceText("MeetingCopper.RibbonMail.xml");
                case "Microsoft.Outlook.Explorer":
                    return GetResourceText("MeetingCopper.Ribbon.xml");
                default:
                    return GetResourceText("MeetingCopper.Ribbon.xml");
            }
        }

        #endregion

        #region Devoluciones de llamada de la cinta de opciones

        //Cree métodos de devolución de llamada aquí. Para obtener más información sobre la adición de métodos de devolución de llamada, visite https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUi) {
            _ribbon = ribbonUi;
        }

        #endregion

        #region Asistentes

        private static string GetResourceText(string resourceName) {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();

            for (int i = 0; i < resourceNames.Length; ++i) {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0) {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i]))) {
                        if (resourceReader != null) {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }

            return null;
        }

        #endregion

        #region Eventos

        /// <summary>
        ///     Evento que crea u modifica una Nueva Reunion
        /// </summary>
        /// <param name="control"></param>
        public void NuevaReunion(Office.IRibbonControl control) {
            Application app = Globals.ThisAddIn.Application;
            AppointmentItem appointment = GetAppointmentItemFromApp(app);

            if (appointment != null) {
                appointment.MeetingStatus = OlMeetingStatus.olMeeting;
                appointment.Start = DateTime.Now.AddHours(2);
                appointment.End = DateTime.Now.AddHours(3);
                appointment.Location = "Elija la ubicación de la Reunión";
                appointment.AllDayEvent = false;
                appointment.Display(false);

                //  Lectura de RTF
                RichTextBox rtb = new RichTextBox {Rtf = Encoding.UTF8.GetString(appointment.RTFBody)};
                rtb.Select(rtb.TextLength, 0);
                rtb.LoadFile(GetTemplateFile(TemplateEnum.NuevaReunion));
                appointment.RTFBody = Encoding.UTF8.GetBytes(rtb.Rtf);
            }
        }

        /// <summary>
        ///     Evento que crea o modifica una nueva Reunión
        /// </summary>
        /// <param name="control"></param>
        public void NuevaRutina(Office.IRibbonControl control) {
            Application app = Globals.ThisAddIn.Application;
            AppointmentItem appointment = GetAppointmentItemFromApp(app);

            if (appointment != null) {
                appointment.MeetingStatus = OlMeetingStatus.olMeeting;
                appointment.Start = DateTime.Now.AddHours(2);
                appointment.End = DateTime.Now.AddHours(3);
                appointment.Location = "Elija la ubicación de la Reunión";
                appointment.AllDayEvent = false;
                appointment.Display(false);

                //  Lectura de RTF
                RichTextBox rtb = new RichTextBox {Rtf = Encoding.UTF8.GetString(appointment.RTFBody)};
                rtb.Select(rtb.TextLength, 0);
                rtb.LoadFile(GetTemplateFile(TemplateEnum.TaskAssigment));
                appointment.RTFBody = Encoding.UTF8.GetBytes(rtb.Rtf);

                //RichTextBox rtbTemp = new RichTextBox();
                //WebBrowser wb = new WebBrowser();
                //wb.Navigate("about:blank");

                //wb.Document.Write(GetSignatureCurrentExchangeUser(app.Session.CurrentUser.AddressEntry.GetExchangeUser()));
                //wb.Document.ExecCommand("SelectAll", false, null);
                //wb.Document.ExecCommand("Copy", false, null);

                //rtbTemp.SelectAll();
                //rtbTemp.Paste();

                //rtb.Rtf = rtb.Rtf + rtbTemp.Rtf;

                //newCita.RTFBody = Encoding.UTF8.GetBytes(rtbTemp.Rtf);
            }
        }

        /// <summary>
        ///     Evento que gerera una nueva minuta
        /// </summary>
        /// <param name="control"></param>
        public void NuevaMinuta(Office.IRibbonControl control) {
            Application app = Globals.ThisAddIn.Application;
            MailItem mail = GetMailItemFromApp(app);

            if (mail != null) {
                string template = LoadTemplate(TemplateEnum.NuevaMinuta);
                string signature = GetSignatureCurrentExchangeUser(app.Session.CurrentUser.AddressEntry.GetExchangeUser());
                mail.Subject = "Minuta Reunión";
                mail.HTMLBody = template + signature;
                mail.Display(false);
            }
        }

        #endregion

        #region Metodos Privados

        #region Iconos

        public Bitmap MinutaIcon(Office.IRibbonControl control) {
            return Resources.minuta;
        }

        public Bitmap RutinaIcon(Office.IRibbonControl control) {
            return Resources.rutina;
        }

        public Bitmap MeetingIcon(Office.IRibbonControl control) {
            return Resources.reunion;
        }

        #endregion

        /// <summary>
        ///     Creo u obtiene el AppointmentItem from App
        /// </summary>
        /// <param name="app"></param>
        /// <returns>AppointmentItem</returns>
        private AppointmentItem GetAppointmentItemFromApp(Application app) {
            if (app.ActiveInspector() != null && app.ActiveInspector().CurrentItem is AppointmentItem) {
                return app.ActiveInspector().CurrentItem as AppointmentItem;
            }

            return (AppointmentItem) app.CreateItem(OlItemType.olAppointmentItem);
        }
        
        
        /// <summary>
        ///     Creo u obtiene el MailItem from App
        /// </summary>
        /// <param name="app"></param>
        /// <returns>MailItem</returns>
        private MailItem GetMailItemFromApp(Application app) {
            if (app.ActiveInspector() != null && app.ActiveInspector().CurrentItem is MailItem) {
                return app.ActiveInspector().CurrentItem as MailItem;
            }

            return (MailItem) app.CreateItem(OlItemType.olMailItem);
        }
        

        /// <summary>
        ///     Lee la firma del usuario segun la cuenta de correo
        /// </summary>
        /// <param name="exchangeUser">Cuenta de correo</param>
        /// <returns>firma en caso de encontrarse</returns>
        private string GetSignatureCurrentExchangeUser(ExchangeUser exchangeUser) {
            string signature = string.Empty;

            if (exchangeUser == null) {
                return signature;
            }

            if (string.IsNullOrEmpty(exchangeUser.PrimarySmtpAddress)) {
                return signature;
            }

            string signatureDirectory = GetSignatureDirectory();

            if (!Directory.Exists(signatureDirectory)) {
                return signature;
            }

            DirectoryInfo di = new DirectoryInfo(signatureDirectory);
            FileInfo[] htmFiles = di.GetFiles("*.htm");

            if (!htmFiles.Any()) {
                return signature;
            }

            string domain = GetDomainFromEmail(exchangeUser.PrimarySmtpAddress);
            List<FileInfo> signatureFiles = htmFiles.Where(f => f.Name.ToLowerInvariant().Contains(domain.ToLowerInvariant())).ToList();

            string signatureFile = signatureFiles.Any() && signatureFiles.Count == 1 ? signatureFiles[0].FullName : htmFiles[0].FullName;

            signature = File.ReadAllText(signatureFile, Encoding.Default);

            return signature;
        }

         /// <summary>
        ///     Metodo que Lee un Template segun el tipo indicado
        /// </summary>
        /// <param name="templateType">Tipo de Template</param>
        /// <returns>Contenido del Template</returns>
        private string LoadTemplate(TemplateEnum templateType) {
            string templateFile = GetTemplateFile(templateType);

            return !File.Exists(templateFile) ? string.Empty : File.ReadAllText(templateFile, Encoding.Default);
        }

        /// <summary>
        ///     Metodo que obtiene la la ruta de la plantilla segun el tipo indicado
        /// </summary>
        /// <param name="templateType">Tipo de Template</param>
        /// <returns>Ruta del archivo</returns>
        private static string GetTemplateFile(TemplateEnum templateType) {
            string templateFile = string.Empty;

            string signatureDirectory = GetTemplateDirectory();

            switch (templateType) {
                case TemplateEnum.NuevaMinuta:
                    templateFile = Path.Combine(signatureDirectory, "MC_NuevaMinuta.html");

                    break;

                case TemplateEnum.NuevaReunion:
                    templateFile = Path.Combine(signatureDirectory, "MC_NuevaReunion.rtf");

                    break;

                case TemplateEnum.TaskAssigment:
                    templateFile = Path.Combine(signatureDirectory, "MC_NuevaRutina.rtf");

                    break;
            }

            return templateFile;
        }

        /// <summary>
        ///     Metodo privado que extrae el dominio de una cuenta de correo
        /// </summary>
        /// <param name="email">correo</param>
        /// <returns>dominio</returns>
        private static string GetDomainFromEmail(string email) {
            try {
                int indexOfAt = email.IndexOf('@');

                return Path.GetFileNameWithoutExtension(email.Substring(indexOfAt + 1));
            } catch {
                return string.Empty;
            }
        }

        /// <summary>
        ///     Obtiene el Directorio de las Firmas del Usuario
        /// </summary>
        /// <returns>Directorio de Firmas del Usuario</returns>
        private static string GetSignatureDirectory() {
            string path = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\Microsoft\\Firmas";

            if (!Directory.Exists(path)) {
                path = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\Microsoft\\Signatures";
            }

            return path;
        }

        /// <summary>
        ///     Obtiene el Directorio de las Plantillas del Usuario
        /// </summary>
        /// <returns>Directorio de Plantillas del Usuario</returns>
        private static string GetTemplateDirectory() {
            string path = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\Microsoft\\Plantillas";

            if (!Directory.Exists(path)) {
                path = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\Microsoft\\Templates";
            }

            return path;
        }
        
        #endregion

            #region OldMethods

        private string ReadSignature(string email) {
            string appDataDir;

            try {
                appDataDir = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\Microsoft\\Signatures";
            } catch {
                appDataDir = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\Microsoft\\Firmas";
            }

            string signature = string.Empty;
            DirectoryInfo diInfo = new DirectoryInfo(appDataDir);

            if (diInfo.Exists) {
                FileInfo[] fiSignature = diInfo.GetFiles("*" + email + "*.htm");

                if (fiSignature.Length > 0) {
                    StreamReader sr = new StreamReader(fiSignature[0].FullName, Encoding.Default);
                    signature = sr.ReadToEnd();

                    if (!string.IsNullOrEmpty(signature)) {
                        string fileName = fiSignature[0].Name.Replace(fiSignature[0].Extension, string.Empty);
                        signature = signature.Replace(fileName + "_files/", appDataDir + "/" + fileName + "_files/");
                    }
                }
            }

            return signature;
        }

        /* public void NuevaCitaBody(Office.IRibbonControl control)
  {
      try
      {
          DialogResult dialogResult = MessageBox.Show("Va a eliminar el contenido no enviado, ¿desea continuar?", "Atención", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
          if (dialogResult == DialogResult.Yes)
          {
              Microsoft.Office.Interop.Outlook.Application app = Globals.ThisAddIn.Application;
              Microsoft.Office.Interop.Outlook.AppointmentItem newCita = (Microsoft.Office.Interop.Outlook.AppointmentItem)
              app.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olAppointmentItem);
              Outlook.Inspector currentInspector = (Outlook.Inspector)newCita.GetInspector;
              Inspectors_NewInspectorv(currentInspector);
          }

      }
      catch (Exception ex)
      {
          //  MessageBox.Show("Oops, ha ocurrido el siguiente error:  " + ex.Message);
      }
  }*/

        /*public void NuevaRutinaBody(Office.IRibbonControl control)
        {
            try
            {
                DialogResult dialogResult = MessageBox.Show("Va a eliminar el contenido no enviado, ¿desea continuar?", "Atención", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (dialogResult == DialogResult.Yes)
                {
                    Microsoft.Office.Interop.Outlook.Application app = Globals.ThisAddIn.Application;
                    Microsoft.Office.Interop.Outlook.AppointmentItem newCita = (Microsoft.Office.Interop.Outlook.AppointmentItem)
                    app.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olAppointmentItem);
                    Outlook.Inspector currentInspector = (Outlook.Inspector)newCita.GetInspector;
                    Inspectors_NewInspectorv(currentInspector);


                }
            }
            catch(Exception ex)
            {

            }
            
        }*/

        /* void Inspectors_NewInspectorv(Outlook.Inspector Inspector)
         {
             Outlook.MailItem mailItem = Inspector.CurrentItem as Outlook.MailItem;
             Outlook.AppointmentItem meetingItem = Inspector.CurrentItem as Outlook.AppointmentItem;
             MessageBox.Show("si");
             if (mailItem != null)
             {
                 MessageBox.Show(mailItem.HTMLBody.Substring(mailItem.HTMLBody.Length - 700));
 
             }
 
             if (meetingItem != null)
             {
                 
             }
         }*/

        /*public void ReadMail(Office.IRibbonControl control)
        {
            try
            {
                
                DialogResult dialogResult = MessageBox.Show("Va a eliminar el contenido no enviado, ¿desea continuar?", "Atención", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (dialogResult == DialogResult.Yes)
                {
                    Outlook.Application app = Globals.ThisAddIn.Application;
                    Outlook.MailItem newMail = (Outlook.MailItem)app.CreateItem(Outlook.OlItemType.olMailItem);
                    Outlook.Inspector currentInspector = (Outlook.Inspector)newMail.GetInspector;
                    Inspectors_NewInspectorv(currentInspector);
                    Outlook.MailItem mailItem = Inspector.CurrentItem as Outlook.MailItem;
                    //Outlook.Inspectors inspectors;
                    //inspectors = this.Application.Inspectors;

                    //Outlook.Application app = Globals.ThisAddIn.Application;
                    //inspectors = app.Inspectors;
                    //Inspectors_NewInspector(inspectors);
                    //inspectors.NewInspector += new Outlook.InspectorsEvents_NewInspectorEventHandler(Inspectors_NewInspectorv);

                    /*Outlook.MailItem newMail = app.CreateItem(Outlook.OlItemType.olMailItem);
                    
                    Outlook.Inspector Inspector = newMail.GetInspector;
                    

                    //Outlook.MailItem newMail = app.CreateItem(Outlook.OlItemType.olMailItem);

                    

                    //Outlook.Inspector inspector = currentInspector;
                    //Outlook.Inspector inspector = app.ActiveInspector.currentInspector;

                    if (Inspector.IsWordMail())
                    {
                        //MessageBox.Show("si");
                        newMail.BodyFormat = Outlook.OlBodyFormat.olFormatHTML;
                        string cuerpo = newMail.HTMLBody;
                        Word.Document wordDocument = Inspector.WordEditor;
                        Word.Selection selected = wordDocument.Windows[1].Selection;
                        MessageBox.Show(cuerpo.Substring(cuerpo.Length - 700));
                        //Word.Range range = wordDocument.Range(wordDocument.Content.Start, wordDocument.Content.End);
                        Word.Range range = wordDocument.Range(0, wordDocument.Characters.Count);
                        //MessageBox.Show(wordDocument.Content.End.ToString());
                        Word.Find findObject = range.Find;
//                        MessageBox.Show(range.Find.);
                        findObject.ClearFormatting();
                        findObject.Text = cuerpo;
                        
                        findObject.Replacement.ClearFormatting();
                        findObject.Replacement.Text = "new value";
                        object replaceAll = Word.WdReplace.wdReplaceAll;

                        findObject.Execute(ReplaceWith: replaceAll);
                    }

                }
            }
            catch(Exception ex)
            {

            }
        }*/

        #endregion

    }

}
