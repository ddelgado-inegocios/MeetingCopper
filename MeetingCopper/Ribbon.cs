using System;
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
        public Bitmap MeetingIcon00(Microsoft.Office.Core.IRibbonControl control) => Resources.meeting2;
        public Bitmap MeetingIcon01(Microsoft.Office.Core.IRibbonControl control) => Resources.meeting01;
        public Bitmap MeetingIcon02(Microsoft.Office.Core.IRibbonControl control) => Resources.meeting02;
        public Bitmap MeetingIcon03(Microsoft.Office.Core.IRibbonControl control) => Resources.meeting03;
        public Bitmap MinutaIcon(Microsoft.Office.Core.IRibbonControl control) => Resources.minuta00;
        public Bitmap RutinaIcon(Microsoft.Office.Core.IRibbonControl control) => Resources.rutina00;

        public void OnClickTema1(Office.IRibbonControl control)
        {
            MessageBox.Show("Tema 1");
        }

        public void OnClickTema2(Office.IRibbonControl control)
        {
            MessageBox.Show("Tema 2");
        }

        public void OnClickTema3(Office.IRibbonControl control)
        {
            MessageBox.Show("Tema 3");
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
