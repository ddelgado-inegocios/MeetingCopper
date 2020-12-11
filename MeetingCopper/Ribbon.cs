using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;
using Word = Microsoft.Office.Interop.Word;
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
        public Bitmap MinutaIcon(Office.IRibbonControl control) => Resources.acuerdo00;
        public Bitmap RutinaIcon(Office.IRibbonControl control) => Resources.rutina00;
        public Bitmap MeetingIcon(Office.IRibbonControl control) => Resources.meeting2;


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
                    try
                    {
                        rtb.LoadFile(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\Microsoft\\Plantillas\\MC_NuevaReunion.rtf");
                    }
                    catch (Exception e)
                    {
                        rtb.LoadFile(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\Microsoft\\Templates\\MC_NuevaReunion.rtf");
                    }
                                        
                    newCita.RTFBody = System.Text.Encoding.UTF8.GetBytes(rtb.Rtf);

                    newCita.MeetingStatus = Microsoft.Office.Interop.Outlook.OlMeetingStatus.olMeeting;
                    
                    newCita.Start = DateTime.Now.AddHours(2);
                    newCita.End = DateTime.Now.AddHours(3);
                    newCita.Location = "Elija la ubicación de la Reunión";
                    //newCita.Subject = "Reunión Template";
                    newCita.Recipients.Add("Seleccione los Destinatarios");
                    newCita.Display(true);
                    newCita.AllDayEvent = false;

                    
                } 
                
            }
            catch (Exception ex)
            {
              //  MessageBox.Show("Oops, ha ocurrido el siguiente error:  " + ex.Message);
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
                    try
                    {
                        rtb.LoadFile(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\Microsoft\\Plantillas\\MC_NuevaRutina.rtf");
                    }
                    catch
                    {
                        rtb.LoadFile(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\Microsoft\\Templates\\MC_NuevaRutina.rtf");
                    }
                    newCita.RTFBody = System.Text.Encoding.UTF8.GetBytes(rtb.Rtf);

                    newCita.Start = DateTime.Now.AddHours(2);
                    newCita.End = DateTime.Now.AddHours(3);
                    newCita.Location = "Elija la ubicación de la Reunión";
                    //newCita.Subject = "Reunión Template";
                    newCita.Recipients.Add("Seleccione los Destinatarios");

                    newCita.Display(true);
                    newCita.AllDayEvent = false;

                    
                }
            }
            catch (Exception ex)
            {
              //  MessageBox.Show("Oops, ha ocurrido el siguiente error:  " + ex.Message);
            }
        }
        
        public void NuevaMinuta(Office.IRibbonControl control)
        {
            
            try
            {
                Microsoft.Office.Interop.Outlook.Application app = Globals.ThisAddIn.Application;
                Microsoft.Office.Interop.Outlook.MailItem newMail = (Microsoft.Office.Interop.Outlook.MailItem)
                app.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);
                Outlook.ExchangeUser currentUser = app.Session.CurrentUser.AddressEntry.GetExchangeUser();
                string email = "@angloamerican";
                if (currentUser != null)
                {
                    email = currentUser.PrimarySmtpAddress;
                }

                if (newMail != null)
                {
                    string HTMLTemplate;
                    try
                    {
                        HTMLTemplate = File.ReadAllText(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\Microsoft\\Plantillas\\MC_NuevaMinuta.html");
                    }
                    catch
                    {
                        HTMLTemplate = File.ReadAllText(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\Microsoft\\Templates\\MC_NuevaMinuta.html");
                    }

                    newMail.Subject = "Minutas de Reunión";
                    newMail.HTMLBody = HTMLTemplate + ReadSignature(email);
                    
                    newMail.To = "Seleccione sus Destinatarios";
                    
                    newMail.Display(true);
                    
                }
            }
            catch (Exception ex)
            {
               // MessageBox.Show("Oops, ha ocurrido el siguiente error:  " + ex.Message);
            }
        }

        private string ReadSignature(string email)
        {
            string appDataDir;
            try
            {
                appDataDir = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\Microsoft\\Firmas";
            }
            catch
            {
                appDataDir = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\Microsoft\\Signatures";
            }

            string signature = string.Empty;
            DirectoryInfo diInfo = new DirectoryInfo(appDataDir);

            if (diInfo.Exists)
            {
                FileInfo[] fiSignature = diInfo.GetFiles("*" + email + "*.htm");

                if (fiSignature.Length > 0)
                {
                    StreamReader sr = new StreamReader(fiSignature[0].FullName, Encoding.Default);
                    signature = sr.ReadToEnd();

                    if (!string.IsNullOrEmpty(signature))
                    {
                        string fileName = fiSignature[0].Name.Replace(fiSignature[0].Extension, string.Empty);
                        signature = signature.Replace(fileName + "_files/", appDataDir + "/" + fileName + "_files/");
                    }
                }
            }
            return signature;
        }

        #region Miembros de IRibbonExtensibility

        public string GetCustomUI(string ribbonID)
        {
            //MessageBox.Show(ribbonID);
            //return GetResourceText("MeetingCopper.Ribbon.xml");
            
            switch (ribbonID)
            {
                /*case "Microsoft.Outlook.Appointment":
                    return GetResourceText("MeetingCopper.RibbonMeeting.xml");
                case "Microsoft.Outlook.Mail.Compose":
                    return GetResourceText("MeetingCopper.RibbonMail.xml");*/
                case "Microsoft.Outlook.Explorer":
                    return GetResourceText("MeetingCopper.Ribbon.xml");
                default:
                    return GetResourceText("");
            }
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
