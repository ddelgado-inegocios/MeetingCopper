// <copyright solution="MeetingTools" project="MeetingTools" file="ThisAddIn.cs">
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
//     File Modified by Cristian Ronny Meneses González at 15-12-2020 1:24
// 
//     Copyright ©2020 Tiaxa Chile. All rights reserved.
// 
// </copyright>
using System;
using Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace MeetingCopper {

    public partial class ThisAddIn {

        private Microsoft.Office.Interop.Outlook.Inspectors _inspectors;

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject() {
            return new Ribbon();
        }

        private void ThisAddIn_Startup(object sender, EventArgs e) {
            _inspectors = Application.Inspectors;
            _inspectors.NewInspector += Inspectors_NewInspector;
        }

        private void Inspectors_NewInspector(Inspector inspector) {
            if (inspector.CurrentItem is AppointmentItem appointment) {
                // appointment.Body = "Este es un cuerpo de una cita "; datos por defecto en cada Cita
            }
            
            if (inspector.CurrentItem is MailItem mail) {
                // mail.Body = "Este es un cuerpo de un correo "; // datos por defecto al crear un correo
            }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e) {
            // Nota: Outlook ya no genera este evento. Si tiene código que 
            //    se debe ejecutar cuando Outlook se apaga, consulte https://go.microsoft.com/fwlink/?LinkId=506785
        }

        #region Código generado por VSTO

        /// <summary>
        ///     Método necesario para admitir el Diseñador. No se puede modificar
        ///     el contenido de este método con el editor de código.
        /// </summary>
        private void InternalStartup() {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }

        #endregion

    }

}

