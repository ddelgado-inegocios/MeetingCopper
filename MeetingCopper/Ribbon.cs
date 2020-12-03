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
        public Bitmap MeetingIcon00(Microsoft.Office.Core.IRibbonControl control) => Resources.meeting2;
        public Bitmap MeetingIcon01(Microsoft.Office.Core.IRibbonControl control) => Resources.meeting01;
        public Bitmap MeetingIcon02(Microsoft.Office.Core.IRibbonControl control) => Resources.meeting02;
        public Bitmap MeetingIcon03(Microsoft.Office.Core.IRibbonControl control) => Resources.meeting03;
        public Bitmap MinutaIcon(Microsoft.Office.Core.IRibbonControl control) => Resources.minuta00;
        public Bitmap RutinaIcon(Microsoft.Office.Core.IRibbonControl control) => Resources.rutina00;


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

                    rtb.SelectedRtf = @"{\rtf1\adeflang1025\ansi\ansicpg1252\uc1\adeff0\deff0\stshfdbch31505\stshfloch31506\stshfhich31506\stshfbi0\deflang3082\deflangfe3082\themelang3082\themelangfe0\themelangcs0{\fonttbl{\f0\fbidi \froman\fcharset0\fprq2{\*\panose 00000000000000000000}Times New Roman;}{\f1\fbidi \fswiss\fcharset0\fprq2{\*\panose 00000000000000000000}Arial;}
{\f34\fbidi \froman\fcharset0\fprq2{\*\panose 02040503050406030204}Cambria Math;}{\f38\fbidi \fswiss\fcharset0\fprq2{\*\panose 00000000000000000000}Calibri Light;}
{\flomajor\f31500\fbidi \froman\fcharset0\fprq2{\*\panose 00000000000000000000}Times New Roman;}{\fdbmajor\f31501\fbidi \froman\fcharset0\fprq2{\*\panose 00000000000000000000}Times New Roman;}
{\fhimajor\f31502\fbidi \fswiss\fcharset0\fprq2{\*\panose 00000000000000000000}Calibri Light;}{\fbimajor\f31503\fbidi \froman\fcharset0\fprq2{\*\panose 00000000000000000000}Times New Roman;}
{\flominor\f31504\fbidi \froman\fcharset0\fprq2{\*\panose 00000000000000000000}Times New Roman;}{\fdbminor\f31505\fbidi \froman\fcharset0\fprq2{\*\panose 00000000000000000000}Times New Roman;}
{\fhiminor\f31506\fbidi \fswiss\fcharset0\fprq2{\*\panose 00000000000000000000}Calibri;}{\fbiminor\f31507\fbidi \froman\fcharset0\fprq2{\*\panose 00000000000000000000}Times New Roman;}{\f43\fbidi \froman\fcharset238\fprq2 Times New Roman CE;}
{\f44\fbidi \froman\fcharset204\fprq2 Times New Roman Cyr;}{\f46\fbidi \froman\fcharset161\fprq2 Times New Roman Greek;}{\f47\fbidi \froman\fcharset162\fprq2 Times New Roman Tur;}{\f48\fbidi \froman\fcharset177\fprq2 Times New Roman (Hebrew);}
{\f49\fbidi \froman\fcharset178\fprq2 Times New Roman (Arabic);}{\f50\fbidi \froman\fcharset186\fprq2 Times New Roman Baltic;}{\f51\fbidi \froman\fcharset163\fprq2 Times New Roman (Vietnamese);}{\f53\fbidi \fswiss\fcharset238\fprq2 Arial CE;}
{\f54\fbidi \fswiss\fcharset204\fprq2 Arial Cyr;}{\f56\fbidi \fswiss\fcharset161\fprq2 Arial Greek;}{\f57\fbidi \fswiss\fcharset162\fprq2 Arial Tur;}{\f58\fbidi \fswiss\fcharset177\fprq2 Arial (Hebrew);}
{\f59\fbidi \fswiss\fcharset178\fprq2 Arial (Arabic);}{\f60\fbidi \fswiss\fcharset186\fprq2 Arial Baltic;}{\f61\fbidi \fswiss\fcharset163\fprq2 Arial (Vietnamese);}{\f383\fbidi \froman\fcharset238\fprq2 Cambria Math CE;}
{\f384\fbidi \froman\fcharset204\fprq2 Cambria Math Cyr;}{\f386\fbidi \froman\fcharset161\fprq2 Cambria Math Greek;}{\f387\fbidi \froman\fcharset162\fprq2 Cambria Math Tur;}{\f390\fbidi \froman\fcharset186\fprq2 Cambria Math Baltic;}
{\f391\fbidi \froman\fcharset163\fprq2 Cambria Math (Vietnamese);}{\f423\fbidi \fswiss\fcharset238\fprq2 Calibri Light CE;}{\f424\fbidi \fswiss\fcharset204\fprq2 Calibri Light Cyr;}{\f426\fbidi \fswiss\fcharset161\fprq2 Calibri Light Greek;}
{\f427\fbidi \fswiss\fcharset162\fprq2 Calibri Light Tur;}{\f428\fbidi \fswiss\fcharset177\fprq2 Calibri Light (Hebrew);}{\f429\fbidi \fswiss\fcharset178\fprq2 Calibri Light (Arabic);}{\f430\fbidi \fswiss\fcharset186\fprq2 Calibri Light Baltic;}
{\f431\fbidi \fswiss\fcharset163\fprq2 Calibri Light (Vietnamese);}{\flomajor\f31508\fbidi \froman\fcharset238\fprq2 Times New Roman CE;}{\flomajor\f31509\fbidi \froman\fcharset204\fprq2 Times New Roman Cyr;}
{\flomajor\f31511\fbidi \froman\fcharset161\fprq2 Times New Roman Greek;}{\flomajor\f31512\fbidi \froman\fcharset162\fprq2 Times New Roman Tur;}{\flomajor\f31513\fbidi \froman\fcharset177\fprq2 Times New Roman (Hebrew);}
{\flomajor\f31514\fbidi \froman\fcharset178\fprq2 Times New Roman (Arabic);}{\flomajor\f31515\fbidi \froman\fcharset186\fprq2 Times New Roman Baltic;}{\flomajor\f31516\fbidi \froman\fcharset163\fprq2 Times New Roman (Vietnamese);}
{\fdbmajor\f31518\fbidi \froman\fcharset238\fprq2 Times New Roman CE;}{\fdbmajor\f31519\fbidi \froman\fcharset204\fprq2 Times New Roman Cyr;}{\fdbmajor\f31521\fbidi \froman\fcharset161\fprq2 Times New Roman Greek;}
{\fdbmajor\f31522\fbidi \froman\fcharset162\fprq2 Times New Roman Tur;}{\fdbmajor\f31523\fbidi \froman\fcharset177\fprq2 Times New Roman (Hebrew);}{\fdbmajor\f31524\fbidi \froman\fcharset178\fprq2 Times New Roman (Arabic);}
{\fdbmajor\f31525\fbidi \froman\fcharset186\fprq2 Times New Roman Baltic;}{\fdbmajor\f31526\fbidi \froman\fcharset163\fprq2 Times New Roman (Vietnamese);}{\fhimajor\f31528\fbidi \fswiss\fcharset238\fprq2 Calibri Light CE;}
{\fhimajor\f31529\fbidi \fswiss\fcharset204\fprq2 Calibri Light Cyr;}{\fhimajor\f31531\fbidi \fswiss\fcharset161\fprq2 Calibri Light Greek;}{\fhimajor\f31532\fbidi \fswiss\fcharset162\fprq2 Calibri Light Tur;}
{\fhimajor\f31533\fbidi \fswiss\fcharset177\fprq2 Calibri Light (Hebrew);}{\fhimajor\f31534\fbidi \fswiss\fcharset178\fprq2 Calibri Light (Arabic);}{\fhimajor\f31535\fbidi \fswiss\fcharset186\fprq2 Calibri Light Baltic;}
{\fhimajor\f31536\fbidi \fswiss\fcharset163\fprq2 Calibri Light (Vietnamese);}{\fbimajor\f31538\fbidi \froman\fcharset238\fprq2 Times New Roman CE;}{\fbimajor\f31539\fbidi \froman\fcharset204\fprq2 Times New Roman Cyr;}
{\fbimajor\f31541\fbidi \froman\fcharset161\fprq2 Times New Roman Greek;}{\fbimajor\f31542\fbidi \froman\fcharset162\fprq2 Times New Roman Tur;}{\fbimajor\f31543\fbidi \froman\fcharset177\fprq2 Times New Roman (Hebrew);}
{\fbimajor\f31544\fbidi \froman\fcharset178\fprq2 Times New Roman (Arabic);}{\fbimajor\f31545\fbidi \froman\fcharset186\fprq2 Times New Roman Baltic;}{\fbimajor\f31546\fbidi \froman\fcharset163\fprq2 Times New Roman (Vietnamese);}
{\flominor\f31548\fbidi \froman\fcharset238\fprq2 Times New Roman CE;}{\flominor\f31549\fbidi \froman\fcharset204\fprq2 Times New Roman Cyr;}{\flominor\f31551\fbidi \froman\fcharset161\fprq2 Times New Roman Greek;}
{\flominor\f31552\fbidi \froman\fcharset162\fprq2 Times New Roman Tur;}{\flominor\f31553\fbidi \froman\fcharset177\fprq2 Times New Roman (Hebrew);}{\flominor\f31554\fbidi \froman\fcharset178\fprq2 Times New Roman (Arabic);}
{\flominor\f31555\fbidi \froman\fcharset186\fprq2 Times New Roman Baltic;}{\flominor\f31556\fbidi \froman\fcharset163\fprq2 Times New Roman (Vietnamese);}{\fdbminor\f31558\fbidi \froman\fcharset238\fprq2 Times New Roman CE;}
{\fdbminor\f31559\fbidi \froman\fcharset204\fprq2 Times New Roman Cyr;}{\fdbminor\f31561\fbidi \froman\fcharset161\fprq2 Times New Roman Greek;}{\fdbminor\f31562\fbidi \froman\fcharset162\fprq2 Times New Roman Tur;}
{\fdbminor\f31563\fbidi \froman\fcharset177\fprq2 Times New Roman (Hebrew);}{\fdbminor\f31564\fbidi \froman\fcharset178\fprq2 Times New Roman (Arabic);}{\fdbminor\f31565\fbidi \froman\fcharset186\fprq2 Times New Roman Baltic;}
{\fdbminor\f31566\fbidi \froman\fcharset163\fprq2 Times New Roman (Vietnamese);}{\fhiminor\f31568\fbidi \fswiss\fcharset238\fprq2 Calibri CE;}{\fhiminor\f31569\fbidi \fswiss\fcharset204\fprq2 Calibri Cyr;}
{\fhiminor\f31571\fbidi \fswiss\fcharset161\fprq2 Calibri Greek;}{\fhiminor\f31572\fbidi \fswiss\fcharset162\fprq2 Calibri Tur;}{\fhiminor\f31573\fbidi \fswiss\fcharset177\fprq2 Calibri (Hebrew);}
{\fhiminor\f31574\fbidi \fswiss\fcharset178\fprq2 Calibri (Arabic);}{\fhiminor\f31575\fbidi \fswiss\fcharset186\fprq2 Calibri Baltic;}{\fhiminor\f31576\fbidi \fswiss\fcharset163\fprq2 Calibri (Vietnamese);}
{\fbiminor\f31578\fbidi \froman\fcharset238\fprq2 Times New Roman CE;}{\fbiminor\f31579\fbidi \froman\fcharset204\fprq2 Times New Roman Cyr;}{\fbiminor\f31581\fbidi \froman\fcharset161\fprq2 Times New Roman Greek;}
{\fbiminor\f31582\fbidi \froman\fcharset162\fprq2 Times New Roman Tur;}{\fbiminor\f31583\fbidi \froman\fcharset177\fprq2 Times New Roman (Hebrew);}{\fbiminor\f31584\fbidi \froman\fcharset178\fprq2 Times New Roman (Arabic);}
{\fbiminor\f31585\fbidi \froman\fcharset186\fprq2 Times New Roman Baltic;}{\fbiminor\f31586\fbidi \froman\fcharset163\fprq2 Times New Roman (Vietnamese);}}{\colortbl;\red0\green0\blue0;\red0\green0\blue255;\red0\green255\blue255;\red0\green255\blue0;
\red255\green0\blue255;\red255\green0\blue0;\red255\green255\blue0;\red255\green255\blue255;\red0\green0\blue128;\red0\green128\blue128;\red0\green128\blue0;\red128\green0\blue128;\red128\green0\blue0;\red128\green128\blue0;\red128\green128\blue128;
\red192\green192\blue192;\red0\green0\blue0;\red0\green0\blue0;\red0\green32\blue96;\red191\green191\blue191;\red255\green255\blue255;\red0\green112\blue192;\red165\green165\blue165;}{\*\defchp \fs22\loch\af31506\hich\af31506\dbch\af31505 }{\*\defpap 
\ql \li0\ri0\sa160\sl259\slmult1\widctlpar\wrapdefault\aspalpha\aspnum\faauto\adjustright\rin0\lin0\itap0 }\noqfpromote {\stylesheet{\ql \li0\ri0\nowidctlpar\wrapdefault\faauto\rin0\lin0\itap0 \rtlch\fcs1 \af0\afs24\alang1025 \ltrch\fcs0 
\fs24\lang1033\langfe255\loch\f0\hich\af0\dbch\af31505\cgrid\langnp1033\langfenp255 \snext0 \sqformat \spriority0 Normal;}{\s1\ql \li0\ri0\sb240\sa120\keepn\nowidctlpar\wrapdefault\faauto\outlinelevel0\rin0\lin0\itap0 \rtlch\fcs1 \ab\af0\afs32\alang1025 
\ltrch\fcs0 \b\fs32\lang1033\langfe255\loch\f0\hich\af0\dbch\af31505\cgrid\langnp1033\langfenp255 \sbasedon0 \snext0 \slink15 \sqformat heading 1;}{\s2\ql \li0\ri0\sb240\sa120\keepn\nowidctlpar\wrapdefault\faauto\outlinelevel1\rin0\lin0\itap0 \rtlch\fcs1 
\ab\ai\af0\afs28\alang1025 \ltrch\fcs0 \b\i\fs28\lang1033\langfe255\loch\f0\hich\af0\dbch\af31505\cgrid\langnp1033\langfenp255 \sbasedon0 \snext0 \slink16 \sqformat heading 2;}{
\s3\ql \li0\ri0\sb240\sa120\keepn\nowidctlpar\wrapdefault\faauto\outlinelevel2\rin0\lin0\itap0 \rtlch\fcs1 \ab\af0\afs28\alang1025 \ltrch\fcs0 \b\fs28\lang1033\langfe255\loch\f0\hich\af0\dbch\af31505\cgrid\langnp1033\langfenp255 
\sbasedon0 \snext0 \slink17 \sqformat heading 3;}{\s4\ql \li0\ri0\sb240\sa120\keepn\nowidctlpar\wrapdefault\faauto\outlinelevel3\rin0\lin0\itap0 \rtlch\fcs1 \ab\ai\af0\afs23\alang1025 \ltrch\fcs0 
\b\i\fs23\lang1033\langfe255\loch\f0\hich\af0\dbch\af31505\cgrid\langnp1033\langfenp255 \sbasedon0 \snext0 \slink18 \sqformat heading 4;}{\s5\ql \li0\ri0\sb240\sa120\keepn\nowidctlpar\wrapdefault\faauto\outlinelevel4\rin0\lin0\itap0 \rtlch\fcs1 
\ab\af0\afs23\alang1025 \ltrch\fcs0 \b\fs23\lang1033\langfe255\loch\f0\hich\af0\dbch\af31505\cgrid\langnp1033\langfenp255 \sbasedon0 \snext0 \slink19 \sqformat heading 5;}{
\s6\ql \li0\ri0\sb240\sa120\keepn\nowidctlpar\wrapdefault\faauto\outlinelevel5\rin0\lin0\itap0 \rtlch\fcs1 \ab\af0\afs21\alang1025 \ltrch\fcs0 \b\fs21\lang1033\langfe255\loch\f0\hich\af0\dbch\af31505\cgrid\langnp1033\langfenp255 
\sbasedon0 \snext0 \slink20 \sqformat heading 6;}{\*\cs10 \additive \ssemihidden \sunhideused \spriority1 Default Paragraph Font;}{\*
\ts11\tsrowd\trftsWidthB3\trpaddl108\trpaddr108\trpaddfl3\trpaddft3\trpaddfb3\trpaddfr3\tblind0\tblindtype3\tsvertalt\tsbrdrt\tsbrdrl\tsbrdrb\tsbrdrr\tsbrdrdgl\tsbrdrdgr\tsbrdrh\tsbrdrv \ql \li0\ri0\sa160\sl259\slmult1
\widctlpar\wrapdefault\aspalpha\aspnum\faauto\adjustright\rin0\lin0\itap0 \rtlch\fcs1 \af0\afs22\alang1025 \ltrch\fcs0 \fs22\lang3082\langfe3082\loch\f31506\hich\af31506\dbch\af31505\cgrid\langnp3082\langfenp3082 \snext11 \ssemihidden \sunhideused 
Normal Table;}{\*\cs15 \additive \rtlch\fcs1 \ab\af0\afs32 \ltrch\fcs0 \b\fs32\lang1033\langfe255\kerning32\loch\f31502\hich\af31502\dbch\af31501\langnp1033\langfenp255 \sbasedon10 \slink1 \slocked \spriority9 Heading 1 Char;}{\*\cs16 \additive 
\rtlch\fcs1 \ab\ai\af0\afs28 \ltrch\fcs0 \b\i\fs28\lang1033\langfe255\loch\f31502\hich\af31502\dbch\af31501\langnp1033\langfenp255 \sbasedon10 \slink2 \slocked \ssemihidden \spriority9 Heading 2 Char;}{\*\cs17 \additive \rtlch\fcs1 \ab\af0\afs26 
\ltrch\fcs0 \b\fs26\lang1033\langfe255\loch\f31502\hich\af31502\dbch\af31501\langnp1033\langfenp255 \sbasedon10 \slink3 \slocked \ssemihidden \spriority9 Heading 3 Char;}{\*\cs18 \additive \rtlch\fcs1 \ab\af0\afs28 \ltrch\fcs0 
\b\fs28\lang1033\langfe255\langnp1033\langfenp255 \sbasedon10 \slink4 \slocked \ssemihidden \spriority9 Heading 4 Char;}{\*\cs19 \additive \rtlch\fcs1 \ab\ai\af0\afs26 \ltrch\fcs0 \b\i\fs26\lang1033\langfe255\langnp1033\langfenp255 
\sbasedon10 \slink5 \slocked \ssemihidden \spriority9 Heading 5 Char;}{\*\cs20 \additive \rtlch\fcs1 \ab\af0 \ltrch\fcs0 \b\lang1033\langfe255\langnp1033\langfenp255 \sbasedon10 \slink6 \slocked \ssemihidden \spriority9 Heading 6 Char;}}{\*\pgptbl {\pgp
\ipgp0\itap0\li0\ri0\sb0\sa0}{\pgp\ipgp0\itap0\li0\ri0\sb0\sa0}{\pgp\ipgp0\itap0\li0\ri0\sb0\sa0}{\pgp\ipgp0\itap0\li0\ri0\sb0\sa0}{\pgp\ipgp0\itap0\li0\ri0\sb0\sa0}{\pgp\ipgp0\itap0\li0\ri0\sb0\sa0}{\pgp\ipgp0\itap0\li0\ri0\sb0\sa0}{\pgp\ipgp0\itap0\li0
\ri0\sb0\sa0}{\pgp\ipgp0\itap0\li0\ri0\sb0\sa0}{\pgp\ipgp0\itap0\li0\ri0\sb0\sa0}{\pgp\ipgp0\itap0\li0\ri0\sb0\sa0}}{\*\rsidtbl \rsid2582123\rsid4933805\rsid5529798\rsid5536621\rsid6626730\rsid9728129\rsid11491465\rsid11756537\rsid12210631\rsid12876912
\rsid12997485\rsid13203125\rsid13584511\rsid14173522\rsid14426457\rsid14811498\rsid15217017\rsid16453471}{\mmathPr\mmathFont34\mbrkBin0\mbrkBinSub0\msmallFrac0\mdispDef1\mlMargin0\mrMargin0\mdefJc1\mwrapIndent1440\mintLim0\mnaryLim1}{\info{\title input}
{\author Unknown}{\operator Danilo Delgado}{\creatim\yr2020\mo12\dy3\hr14\min55}{\revtim\yr2020\mo12\dy3\hr16\min21}{\version17}{\edmins57}{\nofpages1}{\nofwords133}{\nofchars732}{\nofcharsws864}{\vern11}}{\*\xmlnstbl {\xmlns1 http://schemas.microsoft.com
/office/word/2003/wordml}}\paperw12240\paperh15840\margl1440\margr1440\margt1440\margb1440\gutter0\ltrsect 
\widowctrl\ftnbj\aenddoc\hyphhotz425\trackmoves0\trackformatting1\donotembedsysfont0\relyonvml0\donotembedlingdata1\grfdocevents0\validatexml0\showplaceholdtext0\ignoremixedcontent0\saveinvalidxml0\showxmlerrors0\horzdoc\dghspace120\dgvspace120
\dghorigin1701\dgvorigin1984\dghshow0\dgvshow3\jcompress\viewkind1\viewscale100\rsidroot9728129 \fet0{\*\wgrffmtfilter 2450}\ilfomacatclnup0\ltrpar \sectd \ltrsect\linex0\sectdefaultcl\sftnbj {\*\pnseclvl1\pnucrm\pnstart1\pnindent720\pnhang {\pntxta .}}
{\*\pnseclvl2\pnucltr\pnstart1\pnindent720\pnhang {\pntxta .}}{\*\pnseclvl3\pndec\pnstart1\pnindent720\pnhang {\pntxta .}}{\*\pnseclvl4\pnlcltr\pnstart1\pnindent720\pnhang {\pntxta )}}{\*\pnseclvl5\pndec\pnstart1\pnindent720\pnhang {\pntxtb (}{\pntxta )}}
{\*\pnseclvl6\pnlcltr\pnstart1\pnindent720\pnhang {\pntxtb (}{\pntxta )}}{\*\pnseclvl7\pnlcrm\pnstart1\pnindent720\pnhang {\pntxtb (}{\pntxta )}}{\*\pnseclvl8\pnlcltr\pnstart1\pnindent720\pnhang {\pntxtb (}{\pntxta )}}{\*\pnseclvl9
\pnlcrm\pnstart1\pnindent720\pnhang {\pntxtb (}{\pntxta )}}\ltrrow\trowd \irow0\irowband0\ltrrow
\ts11\trgaph70\trrh312\trleft10\trftsWidth3\trwWidth10440\trautofit1\trpaddl70\trpaddr70\trpaddfl3\trpaddft3\trpaddfb3\trpaddfr3\tblrsid6626730\tbllkhdrrows\tbllkhdrcols\tbllknocolband\tblind80\tblindtype3 \clvertalc\clbrdrt\brdrs\brdrw20\brdrcf20 
\clbrdrl\brdrs\brdrw20\brdrcf20 \clbrdrb\brdrs\brdrw20\brdrcf20 \clbrdrr\brdrs\brdrw20\brdrcf20 \clcfpat1\clcbpat19\cltxlrtb\clftsWidth3\clwWidth10440\clcbpatraw19\clcfpatraw1\clhidemark \cellx10450\pard\plain \ltrpar
\qc \li0\ri0\widctlpar\intbl\wrapdefault\aspalpha\aspnum\faauto\adjustright\rin0\lin0\pararsid6626730 \rtlch\fcs1 \af0\afs24\alang1025 \ltrch\fcs0 \fs24\lang1033\langfe255\loch\af0\hich\af0\dbch\af31505\cgrid\langnp1033\langfenp255 {\rtlch\fcs1 
\af1\afs22 \ltrch\fcs0 \f1\fs22\cf8\lang3082\langfe3082\langnp3082\langfenp3082\insrsid6626730\charrsid6626730 Gu\'eda Template para Reuniones \cell }\pard \ltrpar\ql \li0\ri0\sa160\sl259\slmult1
\widctlpar\intbl\wrapdefault\aspalpha\aspnum\faauto\adjustright\rin0\lin0 {\rtlch\fcs1 \af0\afs20 \ltrch\fcs0 \fs20\lang3082\langfe3082\dbch\af0\langnp3082\langfenp3082\insrsid6626730\charrsid6626730 \trowd \irow0\irowband0\ltrrow
\ts11\trgaph70\trrh312\trleft10\trftsWidth3\trwWidth10440\trautofit1\trpaddl70\trpaddr70\trpaddfl3\trpaddft3\trpaddfb3\trpaddfr3\tblrsid6626730\tbllkhdrrows\tbllkhdrcols\tbllknocolband\tblind80\tblindtype3 \clvertalc\clbrdrt\brdrs\brdrw20\brdrcf20 
\clbrdrl\brdrs\brdrw20\brdrcf20 \clbrdrb\brdrs\brdrw20\brdrcf20 \clbrdrr\brdrs\brdrw20\brdrcf20 \clcfpat1\clcbpat19\cltxlrtb\clftsWidth3\clwWidth10440\clcbpatraw19\clcfpatraw1\clhidemark \cellx10450\row \ltrrow}\trowd \irow1\irowband1\ltrrow
\ts11\trgaph70\trrh564\trleft10\trftsWidth3\trwWidth10440\trautofit1\trpaddl70\trpaddr70\trpaddfl3\trpaddft3\trpaddfb3\trpaddfr3\tblrsid6626730\tbllkhdrrows\tbllkhdrcols\tbllknocolband\tblind80\tblindtype3 \clvertalc\clbrdrt\brdrnone \clbrdrl
\brdrs\brdrw20\brdrcf20 \clbrdrb\brdrs\brdrw20\brdrcf20 \clbrdrr\brdrs\brdrw20\brdrcf20 \clcfpat1\clcbpat19\cltxlrtb\clftsWidth3\clwWidth2060\clcbpatraw19\clcfpatraw1\clhidemark \cellx2070\clvertalc\clbrdrt\brdrs\brdrw20\brdrcf22 \clbrdrl
\brdrs\brdrw20\brdrcf22 \clbrdrb\brdrs\brdrw20\brdrcf22 \clbrdrr\brdrs\brdrw20\brdrcf22 \cltxlrtb\clftsWidth3\clwWidth8380\clhidemark \cellx10450\pard \ltrpar
\ql \li0\ri0\widctlpar\intbl\wrapdefault\aspalpha\aspnum\faauto\adjustright\rin0\lin0\pararsid6626730 {\rtlch\fcs1 \af1\afs22 \ltrch\fcs0 \f1\fs22\cf8\lang3082\langfe3082\langnp3082\langfenp3082\insrsid6626730\charrsid6626730 Prop\'f3sito y Objetivos
\cell }{\rtlch\fcs1 \af38\afs18 \ltrch\fcs0 \f38\fs18\cf1\lang3082\langfe3082\langnp3082\langfenp3082\insrsid6626730\charrsid6626730 \'bf}{\rtlch\fcs1 \af38\afs18 \ltrch\fcs0 
\f38\fs18\cf1\lang3082\langfe3082\langnp3082\langfenp3082\insrsid6626730\charrsid6626730 Cu\'e1l}{\rtlch\fcs1 \af38\afs18 \ltrch\fcs0 \f38\fs18\cf1\lang3082\langfe3082\langnp3082\langfenp3082\insrsid6626730\charrsid6626730  es el prop\'f3
sito? Definir los Objetivos Claros\cell }\pard \ltrpar\ql \li0\ri0\sa160\sl259\slmult1\widctlpar\intbl\wrapdefault\aspalpha\aspnum\faauto\adjustright\rin0\lin0 {\rtlch\fcs1 \af0\afs20 \ltrch\fcs0 
\fs20\lang3082\langfe3082\dbch\af0\langnp3082\langfenp3082\insrsid6626730\charrsid6626730 \trowd \irow1\irowband1\ltrrow
\ts11\trgaph70\trrh564\trleft10\trftsWidth3\trwWidth10440\trautofit1\trpaddl70\trpaddr70\trpaddfl3\trpaddft3\trpaddfb3\trpaddfr3\tblrsid6626730\tbllkhdrrows\tbllkhdrcols\tbllknocolband\tblind80\tblindtype3 \clvertalc\clbrdrt\brdrnone \clbrdrl
\brdrs\brdrw20\brdrcf20 \clbrdrb\brdrs\brdrw20\brdrcf20 \clbrdrr\brdrs\brdrw20\brdrcf20 \clcfpat1\clcbpat19\cltxlrtb\clftsWidth3\clwWidth2060\clcbpatraw19\clcfpatraw1\clhidemark \cellx2070\clvertalc\clbrdrt\brdrs\brdrw20\brdrcf22 \clbrdrl
\brdrs\brdrw20\brdrcf22 \clbrdrb\brdrs\brdrw20\brdrcf22 \clbrdrr\brdrs\brdrw20\brdrcf22 \cltxlrtb\clftsWidth3\clwWidth8380\clhidemark \cellx10450\row \ltrrow}\trowd \irow2\irowband2\ltrrow
\ts11\trgaph70\trrh564\trleft10\trftsWidth3\trwWidth10440\trautofit1\trpaddl70\trpaddr70\trpaddfl3\trpaddft3\trpaddfb3\trpaddfr3\tblrsid6626730\tbllkhdrrows\tbllkhdrcols\tbllknocolband\tblind80\tblindtype3 \clvertalc\clbrdrt\brdrnone \clbrdrl
\brdrs\brdrw20\brdrcf20 \clbrdrb\brdrs\brdrw20\brdrcf20 \clbrdrr\brdrs\brdrw20\brdrcf20 \clcfpat1\clcbpat19\cltxlrtb\clftsWidth3\clwWidth2060\clcbpatraw19\clcfpatraw1 \cellx2070\clvertalc\clbrdrt\brdrs\brdrw20\brdrcf22 \clbrdrl\brdrs\brdrw20\brdrcf22 
\clbrdrb\brdrs\brdrw20\brdrcf22 \clbrdrr\brdrs\brdrw20\brdrcf22 \cltxlrtb\clftsWidth3\clwWidth8380 \cellx10450\pard \ltrpar\ql \li0\ri0\widctlpar\intbl\wrapdefault\aspalpha\aspnum\faauto\adjustright\rin0\lin0\pararsid6626730 {\rtlch\fcs1 \af1\afs22 
\ltrch\fcs0 \f1\fs22\cf8\lang3082\langfe3082\langnp3082\langfenp3082\insrsid5529798\charrsid6626730 Participantes}{\rtlch\fcs1 \af1\afs22 \ltrch\fcs0 \f1\fs22\cf8\lang3082\langfe3082\langnp3082\langfenp3082\insrsid5529798\charrsid6626730 \cell }{
\rtlch\fcs1 \af38\afs18 \ltrch\fcs0 \f38\fs18\cf1\lang3082\langfe3082\langnp3082\langfenp3082\insrsid5529798 asdasdasdas}{\rtlch\fcs1 \af38\afs18 \ltrch\fcs0 \f38\fs18\cf1\lang3082\langfe3082\langnp3082\langfenp3082\insrsid5529798\charrsid6626730 \cell 
}\pard \ltrpar\ql \li0\ri0\sa160\sl259\slmult1\widctlpar\intbl\wrapdefault\aspalpha\aspnum\faauto\adjustright\rin0\lin0 {\rtlch\fcs1 \af0\afs20 \ltrch\fcs0 \fs20\lang3082\langfe3082\dbch\af0\langnp3082\langfenp3082\insrsid5529798\charrsid6626730 
\trowd \irow2\irowband2\ltrrow\ts11\trgaph70\trrh564\trleft10\trftsWidth3\trwWidth10440\trautofit1\trpaddl70\trpaddr70\trpaddfl3\trpaddft3\trpaddfb3\trpaddfr3\tblrsid6626730\tbllkhdrrows\tbllkhdrcols\tbllknocolband\tblind80\tblindtype3 \clvertalc\clbrdrt
\brdrnone \clbrdrl\brdrs\brdrw20\brdrcf20 \clbrdrb\brdrs\brdrw20\brdrcf20 \clbrdrr\brdrs\brdrw20\brdrcf20 \clcfpat1\clcbpat19\cltxlrtb\clftsWidth3\clwWidth2060\clcbpatraw19\clcfpatraw1 \cellx2070\clvertalc\clbrdrt\brdrs\brdrw20\brdrcf22 \clbrdrl
\brdrs\brdrw20\brdrcf22 \clbrdrb\brdrs\brdrw20\brdrcf22 \clbrdrr\brdrs\brdrw20\brdrcf22 \cltxlrtb\clftsWidth3\clwWidth8380 \cellx10450\row \ltrrow}\trowd \irow3\irowband3\ltrrow
\ts11\trgaph70\trrh288\trleft10\trftsWidth3\trwWidth10440\trautofit1\trpaddl70\trpaddr70\trpaddfl3\trpaddft3\trpaddfb3\trpaddfr3\tblrsid6626730\tbllkhdrrows\tbllkhdrcols\tbllknocolband\tblind80\tblindtype3 \clvmgf\clvertalc\clbrdrt\brdrnone \clbrdrl
\brdrs\brdrw20\brdrcf20 \clbrdrb\brdrs\brdrw20\brdrcf20 \clbrdrr\brdrs\brdrw20\brdrcf22 \clcfpat1\clcbpat19\cltxlrtb\clftsWidth3\clwWidth2060\clcbpatraw19\clcfpatraw1\clhidemark \cellx2070\clvertalc\clbrdrt\brdrnone \clbrdrl\brdrnone \clbrdrb\brdrnone 
\clbrdrr\brdrs\brdrw20\brdrcf22 \cltxlrtb\clftsWidth3\clwWidth8380\clhidemark \cellx10450\pard \ltrpar\ql \li0\ri0\widctlpar\intbl\wrapdefault\aspalpha\aspnum\faauto\adjustright\rin0\lin0\pararsid6626730 {\rtlch\fcs1 \af1\afs22 \ltrch\fcs0 
\f1\fs22\cf8\lang3082\langfe3082\langnp3082\langfenp3082\insrsid6626730\charrsid6626730 Participantes\cell }{\rtlch\fcs1 \af38\afs18 \ltrch\fcs0 \f38\fs18\cf1\lang3082\langfe3082\langnp3082\langfenp3082\insrsid6626730\charrsid6626730 \bullet  }{
\rtlch\fcs1 \af38\afs18 \ltrch\fcs0 \f38\fs18\cf1\lang3082\langfe3082\langnp3082\langfenp3082\insrsid6626730\charrsid6626730 L\'edder}{\rtlch\fcs1 \af38\afs18 \ltrch\fcs0 
\f38\fs18\cf1\lang3082\langfe3082\langnp3082\langfenp3082\insrsid6626730\charrsid6626730  Reuni\'f3n: \cell }\pard \ltrpar\ql \li0\ri0\sa160\sl259\slmult1\widctlpar\intbl\wrapdefault\aspalpha\aspnum\faauto\adjustright\rin0\lin0 {\rtlch\fcs1 \af0\afs20 
\ltrch\fcs0 \fs20\lang3082\langfe3082\dbch\af0\langnp3082\langfenp3082\insrsid6626730\charrsid6626730 \trowd \irow3\irowband3\ltrrow
\ts11\trgaph70\trrh288\trleft10\trftsWidth3\trwWidth10440\trautofit1\trpaddl70\trpaddr70\trpaddfl3\trpaddft3\trpaddfb3\trpaddfr3\tblrsid6626730\tbllkhdrrows\tbllkhdrcols\tbllknocolband\tblind80\tblindtype3 \clvmgf\clvertalc\clbrdrt\brdrnone \clbrdrl
\brdrs\brdrw20\brdrcf20 \clbrdrb\brdrs\brdrw20\brdrcf20 \clbrdrr\brdrs\brdrw20\brdrcf22 \clcfpat1\clcbpat19\cltxlrtb\clftsWidth3\clwWidth2060\clcbpatraw19\clcfpatraw1\clhidemark \cellx2070\clvertalc\clbrdrt\brdrnone \clbrdrl\brdrnone \clbrdrb\brdrnone 
\clbrdrr\brdrs\brdrw20\brdrcf22 \cltxlrtb\clftsWidth3\clwWidth8380\clhidemark \cellx10450\row \ltrrow}\trowd \irow4\irowband4\ltrrow
\ts11\trgaph70\trrh288\trleft10\trftsWidth3\trwWidth10440\trautofit1\trpaddl70\trpaddr70\trpaddfl3\trpaddft3\trpaddfb3\trpaddfr3\tblrsid6626730\tbllkhdrrows\tbllkhdrcols\tbllknocolband\tblind80\tblindtype3 \clvmrg\clvertalc\clbrdrt\brdrnone \clbrdrl
\brdrs\brdrw20\brdrcf20 \clbrdrb\brdrs\brdrw20\brdrcf20 \clbrdrr\brdrs\brdrw20\brdrcf22 \cltxlrtb\clftsWidth3\clwWidth2060\clshdrawnil\clhidemark \cellx2070\clvertalc\clbrdrt\brdrnone \clbrdrl\brdrnone \clbrdrb\brdrnone \clbrdrr\brdrs\brdrw20\brdrcf22 
\cltxlrtb\clftsWidth3\clwWidth8380\clhidemark \cellx10450\pard \ltrpar\ql \li0\ri0\widctlpar\intbl\wrapdefault\aspalpha\aspnum\faauto\adjustright\rin0\lin0\pararsid6626730 {\rtlch\fcs1 \af1\afs22 \ltrch\fcs0 
\f1\fs22\cf8\lang3082\langfe3082\langnp3082\langfenp3082\insrsid6626730\charrsid6626730 \cell }{\rtlch\fcs1 \af38\afs18 \ltrch\fcs0 \f38\fs18\cf1\lang3082\langfe3082\langnp3082\langfenp3082\insrsid6626730\charrsid6626730 \bullet 
 Facilitador (Acciones & Minuta): \cell }\pard \ltrpar\ql \li0\ri0\sa160\sl259\slmult1\widctlpar\intbl\wrapdefault\aspalpha\aspnum\faauto\adjustright\rin0\lin0 {\rtlch\fcs1 \af0\afs20 \ltrch\fcs0 
\fs20\lang3082\langfe3082\dbch\af0\langnp3082\langfenp3082\insrsid6626730\charrsid6626730 \trowd \irow4\irowband4\ltrrow
\ts11\trgaph70\trrh288\trleft10\trftsWidth3\trwWidth10440\trautofit1\trpaddl70\trpaddr70\trpaddfl3\trpaddft3\trpaddfb3\trpaddfr3\tblrsid6626730\tbllkhdrrows\tbllkhdrcols\tbllknocolband\tblind80\tblindtype3 \clvmrg\clvertalc\clbrdrt\brdrnone \clbrdrl
\brdrs\brdrw20\brdrcf20 \clbrdrb\brdrs\brdrw20\brdrcf20 \clbrdrr\brdrs\brdrw20\brdrcf22 \cltxlrtb\clftsWidth3\clwWidth2060\clshdrawnil\clhidemark \cellx2070\clvertalc\clbrdrt\brdrnone \clbrdrl\brdrnone \clbrdrb\brdrnone \clbrdrr\brdrs\brdrw20\brdrcf22 
\cltxlrtb\clftsWidth3\clwWidth8380\clhidemark \cellx10450\row \ltrrow}\trowd \irow5\irowband5\ltrrow
\ts11\trgaph70\trrh1320\trleft10\trftsWidth3\trwWidth10440\trautofit1\trpaddl70\trpaddr70\trpaddfl3\trpaddft3\trpaddfb3\trpaddfr3\tblrsid6626730\tbllkhdrrows\tbllkhdrcols\tbllknocolband\tblind80\tblindtype3 \clvmrg\clvertalc\clbrdrt\brdrnone \clbrdrl
\brdrs\brdrw20\brdrcf20 \clbrdrb\brdrs\brdrw20\brdrcf20 \clbrdrr\brdrs\brdrw20\brdrcf22 \cltxlrtb\clftsWidth3\clwWidth2060\clshdrawnil\clhidemark \cellx2070\clvertalc\clbrdrt\brdrnone \clbrdrl\brdrnone \clbrdrb\brdrs\brdrw20\brdrcf22 \clbrdrr
\brdrs\brdrw20\brdrcf22 \cltxlrtb\clftsWidth3\clwWidth8380\clhidemark \cellx10450\pard \ltrpar\ql \li0\ri0\widctlpar\intbl\wrapdefault\aspalpha\aspnum\faauto\adjustright\rin0\lin0\pararsid6626730 {\rtlch\fcs1 \af1\afs22 \ltrch\fcs0 
\f1\fs22\cf8\lang3082\langfe3082\langnp3082\langfenp3082\insrsid6626730\charrsid6626730 \cell }{\rtlch\fcs1 \ai\af38\afs18 \ltrch\fcs0 \i\f38\fs18\cf15\lang3082\langfe3082\langnp3082\langfenp3082\insrsid6626730\charrsid6626730 \bullet  
Participantes Requeridos: Participantes 1, -Participantes 2, }{\rtlch\fcs1 \ai\af38\afs18 \ltrch\fcs0 \i\f38\fs18\cf15\lang3082\langfe3082\langnp3082\langfenp3082\insrsid6626730\charrsid6626730 Participante}{\rtlch\fcs1 \ai\af38\afs18 \ltrch\fcs0 
\i\f38\fs18\cf15\lang3082\langfe3082\langnp3082\langfenp3082\insrsid6626730\charrsid6626730  3, ...(Recuerda incluir solo a las personas necesarias, el facilitador puede ayudar a compartir la informaci\'f3
n y minuta a las personas que solo necesiten estar informadas)\cell }\pard \ltrpar\ql \li0\ri0\sa160\sl259\slmult1\widctlpar\intbl\wrapdefault\aspalpha\aspnum\faauto\adjustright\rin0\lin0 {\rtlch\fcs1 \af0\afs20 \ltrch\fcs0 
\fs20\lang3082\langfe3082\dbch\af0\langnp3082\langfenp3082\insrsid6626730\charrsid6626730 \trowd \irow5\irowband5\ltrrow
\ts11\trgaph70\trrh1320\trleft10\trftsWidth3\trwWidth10440\trautofit1\trpaddl70\trpaddr70\trpaddfl3\trpaddft3\trpaddfb3\trpaddfr3\tblrsid6626730\tbllkhdrrows\tbllkhdrcols\tbllknocolband\tblind80\tblindtype3 \clvmrg\clvertalc\clbrdrt\brdrnone \clbrdrl
\brdrs\brdrw20\brdrcf20 \clbrdrb\brdrs\brdrw20\brdrcf20 \clbrdrr\brdrs\brdrw20\brdrcf22 \cltxlrtb\clftsWidth3\clwWidth2060\clshdrawnil\clhidemark \cellx2070\clvertalc\clbrdrt\brdrnone \clbrdrl\brdrnone \clbrdrb\brdrs\brdrw20\brdrcf22 \clbrdrr
\brdrs\brdrw20\brdrcf22 \cltxlrtb\clftsWidth3\clwWidth8380\clhidemark \cellx10450\row \ltrrow}\trowd \irow6\irowband6\ltrrow
\ts11\trgaph70\trrh288\trleft10\trftsWidth3\trwWidth10440\trautofit1\trpaddl70\trpaddr70\trpaddfl3\trpaddft3\trpaddfb3\trpaddfr3\tblrsid6626730\tbllkhdrrows\tbllkhdrcols\tbllknocolband\tblind80\tblindtype3 \clvmgf\clvertalc\clbrdrt\brdrnone \clbrdrl
\brdrs\brdrw20\brdrcf20 \clbrdrb\brdrs\brdrw20\brdrcf20 \clbrdrr\brdrs\brdrw20\brdrcf20 \clcfpat1\clcbpat19\cltxlrtb\clftsWidth3\clwWidth2060\clcbpatraw19\clcfpatraw1\clhidemark \cellx2070\clvertalc\clbrdrt\brdrnone \clbrdrl\brdrnone \clbrdrb\brdrnone 
\clbrdrr\brdrs\brdrw20\brdrcf22 \cltxlrtb\clftsWidth3\clwWidth8380\clhidemark \cellx10450\pard \ltrpar\ql \li0\ri0\widctlpar\intbl\wrapdefault\aspalpha\aspnum\faauto\adjustright\rin0\lin0\pararsid6626730 {\rtlch\fcs1 \af1\afs22 \ltrch\fcs0 
\f1\fs22\cf8\lang3082\langfe3082\langnp3082\langfenp3082\insrsid6626730\charrsid6626730 Agenda\cell }{\rtlch\fcs1 \ai\af38\afs18 \ltrch\fcs0 \i\f38\fs18\cf1\lang3082\langfe3082\langnp3082\langfenp3082\insrsid6626730\charrsid6626730 Tem
a 1 (Tiempo 1) - Responsable 1\cell }\pard \ltrpar\ql \li0\ri0\sa160\sl259\slmult1\widctlpar\intbl\wrapdefault\aspalpha\aspnum\faauto\adjustright\rin0\lin0 {\rtlch\fcs1 \af0\afs20 \ltrch\fcs0 
\fs20\lang3082\langfe3082\dbch\af0\langnp3082\langfenp3082\insrsid6626730\charrsid6626730 \trowd \irow6\irowband6\ltrrow
\ts11\trgaph70\trrh288\trleft10\trftsWidth3\trwWidth10440\trautofit1\trpaddl70\trpaddr70\trpaddfl3\trpaddft3\trpaddfb3\trpaddfr3\tblrsid6626730\tbllkhdrrows\tbllkhdrcols\tbllknocolband\tblind80\tblindtype3 \clvmgf\clvertalc\clbrdrt\brdrnone \clbrdrl
\brdrs\brdrw20\brdrcf20 \clbrdrb\brdrs\brdrw20\brdrcf20 \clbrdrr\brdrs\brdrw20\brdrcf20 \clcfpat1\clcbpat19\cltxlrtb\clftsWidth3\clwWidth2060\clcbpatraw19\clcfpatraw1\clhidemark \cellx2070\clvertalc\clbrdrt\brdrnone \clbrdrl\brdrnone \clbrdrb\brdrnone 
\clbrdrr\brdrs\brdrw20\brdrcf22 \cltxlrtb\clftsWidth3\clwWidth8380\clhidemark \cellx10450\row \ltrrow}\trowd \irow7\irowband7\ltrrow
\ts11\trgaph70\trrh288\trleft10\trftsWidth3\trwWidth10440\trautofit1\trpaddl70\trpaddr70\trpaddfl3\trpaddft3\trpaddfb3\trpaddfr3\tblrsid6626730\tbllkhdrrows\tbllkhdrcols\tbllknocolband\tblind80\tblindtype3 \clvmrg\clvertalc\clbrdrt\brdrnone \clbrdrl
\brdrs\brdrw20\brdrcf20 \clbrdrb\brdrs\brdrw20\brdrcf20 \clbrdrr\brdrs\brdrw20\brdrcf20 \cltxlrtb\clftsWidth3\clwWidth2060\clshdrawnil\clhidemark \cellx2070\clvertalc\clbrdrt\brdrnone \clbrdrl\brdrnone \clbrdrb\brdrnone \clbrdrr\brdrs\brdrw20\brdrcf22 
\cltxlrtb\clftsWidth3\clwWidth8380\clhidemark \cellx10450\pard \ltrpar\ql \li0\ri0\widctlpar\intbl\wrapdefault\aspalpha\aspnum\faauto\adjustright\rin0\lin0\pararsid6626730 {\rtlch\fcs1 \af1\afs22 \ltrch\fcs0 
\f1\fs22\cf8\lang3082\langfe3082\langnp3082\langfenp3082\insrsid6626730\charrsid6626730 \cell }{\rtlch\fcs1 \ai\af38\afs18 \ltrch\fcs0 \i\f38\fs18\cf1\lang3082\langfe3082\langnp3082\langfenp3082\insrsid6626730\charrsid6626730 
Tema 2 (Tiempo 2) - Responsable 2\cell }\pard \ltrpar\ql \li0\ri0\sa160\sl259\slmult1\widctlpar\intbl\wrapdefault\aspalpha\aspnum\faauto\adjustright\rin0\lin0 {\rtlch\fcs1 \af0\afs20 \ltrch\fcs0 
\fs20\lang3082\langfe3082\dbch\af0\langnp3082\langfenp3082\insrsid6626730\charrsid6626730 \trowd \irow7\irowband7\ltrrow
\ts11\trgaph70\trrh288\trleft10\trftsWidth3\trwWidth10440\trautofit1\trpaddl70\trpaddr70\trpaddfl3\trpaddft3\trpaddfb3\trpaddfr3\tblrsid6626730\tbllkhdrrows\tbllkhdrcols\tbllknocolband\tblind80\tblindtype3 \clvmrg\clvertalc\clbrdrt\brdrnone \clbrdrl
\brdrs\brdrw20\brdrcf20 \clbrdrb\brdrs\brdrw20\brdrcf20 \clbrdrr\brdrs\brdrw20\brdrcf20 \cltxlrtb\clftsWidth3\clwWidth2060\clshdrawnil\clhidemark \cellx2070\clvertalc\clbrdrt\brdrnone \clbrdrl\brdrnone \clbrdrb\brdrnone \clbrdrr\brdrs\brdrw20\brdrcf22 
\cltxlrtb\clftsWidth3\clwWidth8380\clhidemark \cellx10450\row \ltrrow}\pard \ltrpar\ql \li0\ri0\widctlpar\intbl\wrapdefault\aspalpha\aspnum\faauto\adjustright\rin0\lin0\pararsid6626730 {\rtlch\fcs1 \af1\afs22 \ltrch\fcs0 
\f1\fs22\cf8\lang3082\langfe3082\langnp3082\langfenp3082\insrsid6626730\charrsid6626730 \cell }{\rtlch\fcs1 \ai\af38\afs18 \ltrch\fcs0 \i\f38\fs18\cf1\lang3082\langfe3082\langnp3082\langfenp3082\insrsid6626730\charrsid6626730 
Tema 3 (Tiempo 3) - Responsable 3\cell }\pard \ltrpar\ql \li0\ri0\sa160\sl259\slmult1\widctlpar\intbl\wrapdefault\aspalpha\aspnum\faauto\adjustright\rin0\lin0 {\rtlch\fcs1 \af0\afs20 \ltrch\fcs0 
\fs20\lang3082\langfe3082\dbch\af0\langnp3082\langfenp3082\insrsid6626730\charrsid6626730 \trowd \irow8\irowband8\ltrrow
\ts11\trgaph70\trrh288\trleft10\trftsWidth3\trwWidth10440\trautofit1\trpaddl70\trpaddr70\trpaddfl3\trpaddft3\trpaddfb3\trpaddfr3\tblrsid6626730\tbllkhdrrows\tbllkhdrcols\tbllknocolband\tblind80\tblindtype3 \clvmrg\clvertalc\clbrdrt\brdrnone \clbrdrl
\brdrs\brdrw20\brdrcf20 \clbrdrb\brdrs\brdrw20\brdrcf20 \clbrdrr\brdrs\brdrw20\brdrcf20 \cltxlrtb\clftsWidth3\clwWidth2060\clshdrawnil\clhidemark \cellx2070\clvertalc\clbrdrt\brdrnone \clbrdrl\brdrnone \clbrdrb\brdrnone \clbrdrr\brdrs\brdrw20\brdrcf22 
\cltxlrtb\clftsWidth3\clwWidth8380\clhidemark \cellx10450\row \ltrrow}\trowd \irow9\irowband9\ltrrow
\ts11\trgaph70\trrh300\trleft10\trftsWidth3\trwWidth10440\trautofit1\trpaddl70\trpaddr70\trpaddfl3\trpaddft3\trpaddfb3\trpaddfr3\tblrsid6626730\tbllkhdrrows\tbllkhdrcols\tbllknocolband\tblind80\tblindtype3 \clvmrg\clvertalc\clbrdrt\brdrnone \clbrdrl
\brdrs\brdrw20\brdrcf20 \clbrdrb\brdrs\brdrw20\brdrcf20 \clbrdrr\brdrs\brdrw20\brdrcf20 \cltxlrtb\clftsWidth3\clwWidth2060\clshdrawnil\clhidemark \cellx2070\clvertalc\clbrdrt\brdrnone \clbrdrl\brdrnone \clbrdrb\brdrnone \clbrdrr\brdrs\brdrw20\brdrcf22 
\cltxlrtb\clftsWidth3\clwWidth8380\clhidemark \cellx10450\pard \ltrpar\ql \li0\ri0\widctlpar\intbl\wrapdefault\aspalpha\aspnum\faauto\adjustright\rin0\lin0\pararsid6626730 {\rtlch\fcs1 \af1\afs22 \ltrch\fcs0 
\f1\fs22\cf8\lang3082\langfe3082\langnp3082\langfenp3082\insrsid6626730\charrsid6626730 \cell }{\rtlch\fcs1 \af38\afs18 \ltrch\fcs0 \f38\fs18\cf1\lang3082\langfe3082\langnp3082\langfenp3082\insrsid6626730\charrsid6626730 
(Recuerda dejar explicito si el tema es para informar, discutir, aprobar)\cell }\pard \ltrpar\ql \li0\ri0\sa160\sl259\slmult1\widctlpar\intbl\wrapdefault\aspalpha\aspnum\faauto\adjustright\rin0\lin0 {\rtlch\fcs1 \af0\afs20 \ltrch\fcs0 
\fs20\lang3082\langfe3082\dbch\af0\langnp3082\langfenp3082\insrsid6626730\charrsid6626730 \trowd \irow9\irowband9\ltrrow
\ts11\trgaph70\trrh300\trleft10\trftsWidth3\trwWidth10440\trautofit1\trpaddl70\trpaddr70\trpaddfl3\trpaddft3\trpaddfb3\trpaddfr3\tblrsid6626730\tbllkhdrrows\tbllkhdrcols\tbllknocolband\tblind80\tblindtype3 \clvmrg\clvertalc\clbrdrt\brdrnone \clbrdrl
\brdrs\brdrw20\brdrcf20 \clbrdrb\brdrs\brdrw20\brdrcf20 \clbrdrr\brdrs\brdrw20\brdrcf20 \cltxlrtb\clftsWidth3\clwWidth2060\clshdrawnil\clhidemark \cellx2070\clvertalc\clbrdrt\brdrnone \clbrdrl\brdrnone \clbrdrb\brdrnone \clbrdrr\brdrs\brdrw20\brdrcf22 
\cltxlrtb\clftsWidth3\clwWidth8380\clhidemark \cellx10450\row \ltrrow}\trowd \irow10\irowband10\lastrow \ltrrow
\ts11\trgaph70\trrh918\trleft10\trftsWidth3\trwWidth10440\trautofit1\trpaddl70\trpaddr70\trpaddfl3\trpaddft3\trpaddfb3\trpaddfr3\tblrsid6626730\tbllkhdrrows\tbllkhdrcols\tbllknocolband\tblind80\tblindtype3 \clvertalc\clbrdrt\brdrnone \clbrdrl
\brdrs\brdrw20\brdrcf20 \clbrdrb\brdrs\brdrw20\brdrcf20 \clbrdrr\brdrs\brdrw20\brdrcf20 \clcfpat1\clcbpat19\cltxlrtb\clftsWidth3\clwWidth2060\clcbpatraw19\clcfpatraw1\clhidemark \cellx2070\clvertalc\clbrdrt\brdrs\brdrw20\brdrcf22 \clbrdrl
\brdrs\brdrw20\brdrcf22 \clbrdrb\brdrs\brdrw20\brdrcf22 \clbrdrr\brdrs\brdrw20\brdrcf22 \cltxlrtb\clftsWidth3\clwWidth8380\clhidemark \cellx10450\pard \ltrpar
\ql \li0\ri0\widctlpar\intbl\wrapdefault\aspalpha\aspnum\faauto\adjustright\rin0\lin0\pararsid6626730 {\rtlch\fcs1 \af1\afs22 \ltrch\fcs0 \f1\fs22\cf8\lang3082\langfe3082\langnp3082\langfenp3082\insrsid6626730\charrsid6626730 
Herramientas y/o Material necesario\cell }{\rtlch\fcs1 \ai\af38\afs18 \ltrch\fcs0 \i\f38\fs18\cf1\lang1033\langfe3082\langfenp3082\insrsid6626730\charrsid6626730 Por ejemplo: MS Teams,  SharePoint , Dashboard en PBI, Planner, Menti. }{\rtlch\fcs1 
\ai\af38\afs18 \ltrch\fcs0 \i\f38\fs18\cf23\lang3082\langfe3082\langnp3082\langfenp3082\insrsid6626730\charrsid6626730 Tambi\'e9n}{\rtlch\fcs1 \ai\af38\afs18 \ltrch\fcs0 
\i\f38\fs18\cf23\lang3082\langfe3082\langnp3082\langfenp3082\insrsid6626730\charrsid6626730  puedes indicar si hay material que deba ser revisado con anticipaci\'f3n y asegura de enviarlo con tiempo.\cell }\pard \ltrpar\ql \li0\ri0\sa160\sl259\slmult1
\widctlpar\intbl\wrapdefault\aspalpha\aspnum\faauto\adjustright\rin0\lin0 {\rtlch\fcs1 \af0\afs20 \ltrch\fcs0 \fs20\lang3082\langfe3082\dbch\af0\langnp3082\langfenp3082\insrsid6626730\charrsid6626730 \trowd \irow10\irowband10\lastrow \ltrrow
\ts11\trgaph70\trrh918\trleft10\trftsWidth3\trwWidth10440\trautofit1\trpaddl70\trpaddr70\trpaddfl3\trpaddft3\trpaddfb3\trpaddfr3\tblrsid6626730\tbllkhdrrows\tbllkhdrcols\tbllknocolband\tblind80\tblindtype3 \clvertalc\clbrdrt\brdrnone \clbrdrl
\brdrs\brdrw20\brdrcf20 \clbrdrb\brdrs\brdrw20\brdrcf20 \clbrdrr\brdrs\brdrw20\brdrcf20 \clcfpat1\clcbpat19\cltxlrtb\clftsWidth3\clwWidth2060\clcbpatraw19\clcfpatraw1\clhidemark \cellx2070\clvertalc\clbrdrt\brdrs\brdrw20\brdrcf22 \clbrdrl
\brdrs\brdrw20\brdrcf22 \clbrdrb\brdrs\brdrw20\brdrcf22 \clbrdrr\brdrs\brdrw20\brdrcf22 \cltxlrtb\clftsWidth3\clwWidth8380\clhidemark \cellx10450\row }\pard \ltrpar\ql \li0\ri0\nowidctlpar\wrapdefault\faauto\rin0\lin0\itap0\pararsid12210631 {\rtlch\fcs1 
\af0 \ltrch\fcs0 \lang3082\langfe255\langnp3082\insrsid15217017\charrsid14426457 
\par }";
                    newCita.RTFBody = System.Text.Encoding.UTF8.GetBytes(rtb.Rtf);

                    newCita.MeetingStatus = Microsoft.Office.Interop.Outlook.OlMeetingStatus.olMeeting;
                    
                    newCita.Start = DateTime.Now.AddHours(2);
                    newCita.End = DateTime.Now.AddHours(3);
                    newCita.Location = "Elija la ubicación de la Reunión";
                    newCita.Subject = "Reunión Template";
                    newCita.Recipients.Add("Danilo Delgado Redlich");
                    Microsoft.Office.Interop.Outlook.Recipients sentTo = newCita.Recipients;
                    Microsoft.Office.Interop.Outlook.Recipient sentInvite = null;
                    sentInvite = sentTo.Add("Juanito Perez");
                    sentInvite.Type = (int)Microsoft.Office.Interop.Outlook.OlMeetingRecipientType.olRequired;
                    sentInvite = sentTo.Add("Luis Mario Gonzalez");
                    sentInvite.Type = (int)Microsoft.Office.Interop.Outlook.OlMeetingRecipientType.olOptional;
                    sentTo.ResolveAll();
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

                    rtb.SelectedRtf = @"{\rtf1\adeflang1025\ansi\ansicpg1252\uc1\adeff0\deff0\stshfdbch31505\stshfloch31506\stshfhich31506\stshfbi0\deflang3082\deflangfe3082\themelang3082\themelangfe0\themelangcs0{\fonttbl{\f0\fbidi \froman\fcharset0\fprq2{\*\panose 02020603050405020304}Times New Roman;}{\f1\fbidi \fswiss\fcharset0\fprq2{\*\panose 020b0604020202020204}Arial;}
{\f34\fbidi \froman\fcharset0\fprq2{\*\panose 02040503050406030204}Cambria Math;}{\flomajor\f31500\fbidi \froman\fcharset0\fprq2{\*\panose 02020603050405020304}Times New Roman;}
{\fdbmajor\f31501\fbidi \froman\fcharset0\fprq2{\*\panose 02020603050405020304}Times New Roman;}{\fhimajor\f31502\fbidi \fswiss\fcharset0\fprq2{\*\panose 020f0302020204030204}Calibri Light;}
{\fbimajor\f31503\fbidi \froman\fcharset0\fprq2{\*\panose 02020603050405020304}Times New Roman;}{\flominor\f31504\fbidi \froman\fcharset0\fprq2{\*\panose 02020603050405020304}Times New Roman;}
{\fdbminor\f31505\fbidi \froman\fcharset0\fprq2{\*\panose 02020603050405020304}Times New Roman;}{\fhiminor\f31506\fbidi \fswiss\fcharset0\fprq2{\*\panose 020f0502020204030204}Calibri;}
{\fbiminor\f31507\fbidi \froman\fcharset0\fprq2{\*\panose 02020603050405020304}Times New Roman;}{\f43\fbidi \froman\fcharset238\fprq2 Times New Roman CE;}{\f44\fbidi \froman\fcharset204\fprq2 Times New Roman Cyr;}
{\f46\fbidi \froman\fcharset161\fprq2 Times New Roman Greek;}{\f47\fbidi \froman\fcharset162\fprq2 Times New Roman Tur;}{\f48\fbidi \froman\fcharset177\fprq2 Times New Roman (Hebrew);}{\f49\fbidi \froman\fcharset178\fprq2 Times New Roman (Arabic);}
{\f50\fbidi \froman\fcharset186\fprq2 Times New Roman Baltic;}{\f51\fbidi \froman\fcharset163\fprq2 Times New Roman (Vietnamese);}{\f53\fbidi \fswiss\fcharset238\fprq2 Arial CE;}{\f54\fbidi \fswiss\fcharset204\fprq2 Arial Cyr;}
{\f56\fbidi \fswiss\fcharset161\fprq2 Arial Greek;}{\f57\fbidi \fswiss\fcharset162\fprq2 Arial Tur;}{\f58\fbidi \fswiss\fcharset177\fprq2 Arial (Hebrew);}{\f59\fbidi \fswiss\fcharset178\fprq2 Arial (Arabic);}
{\f60\fbidi \fswiss\fcharset186\fprq2 Arial Baltic;}{\f61\fbidi \fswiss\fcharset163\fprq2 Arial (Vietnamese);}{\f383\fbidi \froman\fcharset238\fprq2 Cambria Math CE;}{\f384\fbidi \froman\fcharset204\fprq2 Cambria Math Cyr;}
{\f386\fbidi \froman\fcharset161\fprq2 Cambria Math Greek;}{\f387\fbidi \froman\fcharset162\fprq2 Cambria Math Tur;}{\f390\fbidi \froman\fcharset186\fprq2 Cambria Math Baltic;}{\f391\fbidi \froman\fcharset163\fprq2 Cambria Math (Vietnamese);}
{\flomajor\f31508\fbidi \froman\fcharset238\fprq2 Times New Roman CE;}{\flomajor\f31509\fbidi \froman\fcharset204\fprq2 Times New Roman Cyr;}{\flomajor\f31511\fbidi \froman\fcharset161\fprq2 Times New Roman Greek;}
{\flomajor\f31512\fbidi \froman\fcharset162\fprq2 Times New Roman Tur;}{\flomajor\f31513\fbidi \froman\fcharset177\fprq2 Times New Roman (Hebrew);}{\flomajor\f31514\fbidi \froman\fcharset178\fprq2 Times New Roman (Arabic);}
{\flomajor\f31515\fbidi \froman\fcharset186\fprq2 Times New Roman Baltic;}{\flomajor\f31516\fbidi \froman\fcharset163\fprq2 Times New Roman (Vietnamese);}{\fdbmajor\f31518\fbidi \froman\fcharset238\fprq2 Times New Roman CE;}
{\fdbmajor\f31519\fbidi \froman\fcharset204\fprq2 Times New Roman Cyr;}{\fdbmajor\f31521\fbidi \froman\fcharset161\fprq2 Times New Roman Greek;}{\fdbmajor\f31522\fbidi \froman\fcharset162\fprq2 Times New Roman Tur;}
{\fdbmajor\f31523\fbidi \froman\fcharset177\fprq2 Times New Roman (Hebrew);}{\fdbmajor\f31524\fbidi \froman\fcharset178\fprq2 Times New Roman (Arabic);}{\fdbmajor\f31525\fbidi \froman\fcharset186\fprq2 Times New Roman Baltic;}
{\fdbmajor\f31526\fbidi \froman\fcharset163\fprq2 Times New Roman (Vietnamese);}{\fhimajor\f31528\fbidi \fswiss\fcharset238\fprq2 Calibri Light CE;}{\fhimajor\f31529\fbidi \fswiss\fcharset204\fprq2 Calibri Light Cyr;}
{\fhimajor\f31531\fbidi \fswiss\fcharset161\fprq2 Calibri Light Greek;}{\fhimajor\f31532\fbidi \fswiss\fcharset162\fprq2 Calibri Light Tur;}{\fhimajor\f31533\fbidi \fswiss\fcharset177\fprq2 Calibri Light (Hebrew);}
{\fhimajor\f31534\fbidi \fswiss\fcharset178\fprq2 Calibri Light (Arabic);}{\fhimajor\f31535\fbidi \fswiss\fcharset186\fprq2 Calibri Light Baltic;}{\fhimajor\f31536\fbidi \fswiss\fcharset163\fprq2 Calibri Light (Vietnamese);}
{\fbimajor\f31538\fbidi \froman\fcharset238\fprq2 Times New Roman CE;}{\fbimajor\f31539\fbidi \froman\fcharset204\fprq2 Times New Roman Cyr;}{\fbimajor\f31541\fbidi \froman\fcharset161\fprq2 Times New Roman Greek;}
{\fbimajor\f31542\fbidi \froman\fcharset162\fprq2 Times New Roman Tur;}{\fbimajor\f31543\fbidi \froman\fcharset177\fprq2 Times New Roman (Hebrew);}{\fbimajor\f31544\fbidi \froman\fcharset178\fprq2 Times New Roman (Arabic);}
{\fbimajor\f31545\fbidi \froman\fcharset186\fprq2 Times New Roman Baltic;}{\fbimajor\f31546\fbidi \froman\fcharset163\fprq2 Times New Roman (Vietnamese);}{\flominor\f31548\fbidi \froman\fcharset238\fprq2 Times New Roman CE;}
{\flominor\f31549\fbidi \froman\fcharset204\fprq2 Times New Roman Cyr;}{\flominor\f31551\fbidi \froman\fcharset161\fprq2 Times New Roman Greek;}{\flominor\f31552\fbidi \froman\fcharset162\fprq2 Times New Roman Tur;}
{\flominor\f31553\fbidi \froman\fcharset177\fprq2 Times New Roman (Hebrew);}{\flominor\f31554\fbidi \froman\fcharset178\fprq2 Times New Roman (Arabic);}{\flominor\f31555\fbidi \froman\fcharset186\fprq2 Times New Roman Baltic;}
{\flominor\f31556\fbidi \froman\fcharset163\fprq2 Times New Roman (Vietnamese);}{\fdbminor\f31558\fbidi \froman\fcharset238\fprq2 Times New Roman CE;}{\fdbminor\f31559\fbidi \froman\fcharset204\fprq2 Times New Roman Cyr;}
{\fdbminor\f31561\fbidi \froman\fcharset161\fprq2 Times New Roman Greek;}{\fdbminor\f31562\fbidi \froman\fcharset162\fprq2 Times New Roman Tur;}{\fdbminor\f31563\fbidi \froman\fcharset177\fprq2 Times New Roman (Hebrew);}
{\fdbminor\f31564\fbidi \froman\fcharset178\fprq2 Times New Roman (Arabic);}{\fdbminor\f31565\fbidi \froman\fcharset186\fprq2 Times New Roman Baltic;}{\fdbminor\f31566\fbidi \froman\fcharset163\fprq2 Times New Roman (Vietnamese);}
{\fhiminor\f31568\fbidi \fswiss\fcharset238\fprq2 Calibri CE;}{\fhiminor\f31569\fbidi \fswiss\fcharset204\fprq2 Calibri Cyr;}{\fhiminor\f31571\fbidi \fswiss\fcharset161\fprq2 Calibri Greek;}{\fhiminor\f31572\fbidi \fswiss\fcharset162\fprq2 Calibri Tur;}
{\fhiminor\f31573\fbidi \fswiss\fcharset177\fprq2 Calibri (Hebrew);}{\fhiminor\f31574\fbidi \fswiss\fcharset178\fprq2 Calibri (Arabic);}{\fhiminor\f31575\fbidi \fswiss\fcharset186\fprq2 Calibri Baltic;}
{\fhiminor\f31576\fbidi \fswiss\fcharset163\fprq2 Calibri (Vietnamese);}{\fbiminor\f31578\fbidi \froman\fcharset238\fprq2 Times New Roman CE;}{\fbiminor\f31579\fbidi \froman\fcharset204\fprq2 Times New Roman Cyr;}
{\fbiminor\f31581\fbidi \froman\fcharset161\fprq2 Times New Roman Greek;}{\fbiminor\f31582\fbidi \froman\fcharset162\fprq2 Times New Roman Tur;}{\fbiminor\f31583\fbidi \froman\fcharset177\fprq2 Times New Roman (Hebrew);}
{\fbiminor\f31584\fbidi \froman\fcharset178\fprq2 Times New Roman (Arabic);}{\fbiminor\f31585\fbidi \froman\fcharset186\fprq2 Times New Roman Baltic;}{\fbiminor\f31586\fbidi \froman\fcharset163\fprq2 Times New Roman (Vietnamese);}}
{\colortbl;\red0\green0\blue0;\red0\green0\blue255;\red0\green255\blue255;\red0\green255\blue0;\red255\green0\blue255;\red255\green0\blue0;\red255\green255\blue0;\red255\green255\blue255;\red0\green0\blue128;\red0\green128\blue128;\red0\green128\blue0;
\red128\green0\blue128;\red128\green0\blue0;\red128\green128\blue0;\red128\green128\blue128;\red192\green192\blue192;\red0\green0\blue0;\red0\green0\blue0;\red0\green32\blue96;\red191\green191\blue191;\red255\green255\blue255;}{\*\defchp 
\fs22\loch\af31506\hich\af31506\dbch\af31505 }{\*\defpap \ql \li0\ri0\sa160\sl259\slmult1\widctlpar\wrapdefault\aspalpha\aspnum\faauto\adjustright\rin0\lin0\itap0 }\noqfpromote {\stylesheet{\ql \li0\ri0\nowidctlpar\wrapdefault\faauto\rin0\lin0\itap0 
\rtlch\fcs1 \af0\afs24\alang1025 \ltrch\fcs0 \fs24\lang1033\langfe255\loch\f0\hich\af0\dbch\af31505\cgrid\langnp1033\langfenp255 \snext0 \sqformat \spriority0 Normal;}{
\s1\ql \li0\ri0\sb240\sa120\keepn\nowidctlpar\wrapdefault\faauto\outlinelevel0\rin0\lin0\itap0 \rtlch\fcs1 \ab\af0\afs32\alang1025 \ltrch\fcs0 \b\fs32\lang1033\langfe255\loch\f0\hich\af0\dbch\af31505\cgrid\langnp1033\langfenp255 
\sbasedon0 \snext0 \slink15 \sqformat heading 1;}{\s2\ql \li0\ri0\sb240\sa120\keepn\nowidctlpar\wrapdefault\faauto\outlinelevel1\rin0\lin0\itap0 \rtlch\fcs1 \ab\ai\af0\afs28\alang1025 \ltrch\fcs0 
\b\i\fs28\lang1033\langfe255\loch\f0\hich\af0\dbch\af31505\cgrid\langnp1033\langfenp255 \sbasedon0 \snext0 \slink16 \sqformat heading 2;}{\s3\ql \li0\ri0\sb240\sa120\keepn\nowidctlpar\wrapdefault\faauto\outlinelevel2\rin0\lin0\itap0 \rtlch\fcs1 
\ab\af0\afs28\alang1025 \ltrch\fcs0 \b\fs28\lang1033\langfe255\loch\f0\hich\af0\dbch\af31505\cgrid\langnp1033\langfenp255 \sbasedon0 \snext0 \slink17 \sqformat heading 3;}{
\s4\ql \li0\ri0\sb240\sa120\keepn\nowidctlpar\wrapdefault\faauto\outlinelevel3\rin0\lin0\itap0 \rtlch\fcs1 \ab\ai\af0\afs23\alang1025 \ltrch\fcs0 \b\i\fs23\lang1033\langfe255\loch\f0\hich\af0\dbch\af31505\cgrid\langnp1033\langfenp255 
\sbasedon0 \snext0 \slink18 \sqformat heading 4;}{\s5\ql \li0\ri0\sb240\sa120\keepn\nowidctlpar\wrapdefault\faauto\outlinelevel4\rin0\lin0\itap0 \rtlch\fcs1 \ab\af0\afs23\alang1025 \ltrch\fcs0 
\b\fs23\lang1033\langfe255\loch\f0\hich\af0\dbch\af31505\cgrid\langnp1033\langfenp255 \sbasedon0 \snext0 \slink19 \sqformat heading 5;}{\s6\ql \li0\ri0\sb240\sa120\keepn\nowidctlpar\wrapdefault\faauto\outlinelevel5\rin0\lin0\itap0 \rtlch\fcs1 
\ab\af0\afs21\alang1025 \ltrch\fcs0 \b\fs21\lang1033\langfe255\loch\f0\hich\af0\dbch\af31505\cgrid\langnp1033\langfenp255 \sbasedon0 \snext0 \slink20 \sqformat heading 6;}{\*\cs10 \additive \ssemihidden \sunhideused \spriority1 Default Paragraph Font;}{\*
\ts11\tsrowd\trftsWidthB3\trpaddl108\trpaddr108\trpaddfl3\trpaddft3\trpaddfb3\trpaddfr3\tblind0\tblindtype3\tsvertalt\tsbrdrt\tsbrdrl\tsbrdrb\tsbrdrr\tsbrdrdgl\tsbrdrdgr\tsbrdrh\tsbrdrv \ql \li0\ri0\sa160\sl259\slmult1
\widctlpar\wrapdefault\aspalpha\aspnum\faauto\adjustright\rin0\lin0\itap0 \rtlch\fcs1 \af0\afs22\alang1025 \ltrch\fcs0 \fs22\lang3082\langfe3082\loch\f31506\hich\af31506\dbch\af31505\cgrid\langnp3082\langfenp3082 \snext11 \ssemihidden \sunhideused 
Normal Table;}{\*\cs15 \additive \rtlch\fcs1 \ab\af0\afs32 \ltrch\fcs0 \b\fs32\lang1033\langfe255\kerning32\loch\f31502\hich\af31502\dbch\af31501\langnp1033\langfenp255 \sbasedon10 \slink1 \slocked \spriority9 Heading 1 Char;}{\*\cs16 \additive 
\rtlch\fcs1 \ab\ai\af0\afs28 \ltrch\fcs0 \b\i\fs28\lang1033\langfe255\loch\f31502\hich\af31502\dbch\af31501\langnp1033\langfenp255 \sbasedon10 \slink2 \slocked \ssemihidden \spriority9 Heading 2 Char;}{\*\cs17 \additive \rtlch\fcs1 \ab\af0\afs26 
\ltrch\fcs0 \b\fs26\lang1033\langfe255\loch\f31502\hich\af31502\dbch\af31501\langnp1033\langfenp255 \sbasedon10 \slink3 \slocked \ssemihidden \spriority9 Heading 3 Char;}{\*\cs18 \additive \rtlch\fcs1 \ab\af0\afs28 \ltrch\fcs0 
\b\fs28\lang1033\langfe255\langnp1033\langfenp255 \sbasedon10 \slink4 \slocked \ssemihidden \spriority9 Heading 4 Char;}{\*\cs19 \additive \rtlch\fcs1 \ab\ai\af0\afs26 \ltrch\fcs0 \b\i\fs26\lang1033\langfe255\langnp1033\langfenp255 
\sbasedon10 \slink5 \slocked \ssemihidden \spriority9 Heading 5 Char;}{\*\cs20 \additive \rtlch\fcs1 \ab\af0 \ltrch\fcs0 \b\lang1033\langfe255\langnp1033\langfenp255 \sbasedon10 \slink6 \slocked \ssemihidden \spriority9 Heading 6 Char;}}{\*\pgptbl {\pgp
\ipgp0\itap0\li0\ri0\sb0\sa0}{\pgp\ipgp0\itap0\li0\ri0\sb0\sa0}{\pgp\ipgp0\itap0\li0\ri0\sb0\sa0}{\pgp\ipgp0\itap0\li0\ri0\sb0\sa0}{\pgp\ipgp0\itap0\li0\ri0\sb0\sa0}{\pgp\ipgp0\itap0\li0\ri0\sb0\sa0}{\pgp\ipgp0\itap0\li0\ri0\sb0\sa0}{\pgp\ipgp0\itap0\li0
\ri0\sb0\sa0}{\pgp\ipgp0\itap0\li0\ri0\sb0\sa0}}{\*\rsidtbl \rsid2582123\rsid4933805\rsid5536621\rsid9728129\rsid11491465\rsid12210631\rsid12997485\rsid13584511\rsid14173522\rsid14426457\rsid15217017}{\mmathPr\mmathFont34\mbrkBin0\mbrkBinSub0\msmallFrac0
\mdispDef1\mlMargin0\mrMargin0\mdefJc1\mwrapIndent1440\mintLim0\mnaryLim1}{\info{\title input}{\author Unknown}{\operator Danilo Delgado}{\creatim\yr2020\mo12\dy3\hr14\min55}{\revtim\yr2020\mo12\dy3\hr15\min40}{\version10}{\edmins16}{\nofpages1}
{\nofwords82}{\nofchars455}{\nofcharsws536}{\vern11}}{\*\xmlnstbl {\xmlns1 http://schemas.microsoft.com/office/word/2003/wordml}}\paperw12240\paperh15840\margl1440\margr1440\margt1440\margb1440\gutter0\ltrsect 
\widowctrl\ftnbj\aenddoc\hyphhotz425\trackmoves0\trackformatting1\donotembedsysfont0\relyonvml0\donotembedlingdata1\grfdocevents0\validatexml0\showplaceholdtext0\ignoremixedcontent0\saveinvalidxml0\showxmlerrors0\horzdoc\dghspace120\dgvspace120
\dghorigin1701\dgvorigin1984\dghshow0\dgvshow3\jcompress\viewkind1\viewscale100\rsidroot9728129 \fet0{\*\wgrffmtfilter 2450}\ilfomacatclnup0\ltrpar \sectd \ltrsect\linex0\sectdefaultcl\sftnbj {\*\pnseclvl1\pnucrm\pnstart1\pnindent720\pnhang {\pntxta .}}
{\*\pnseclvl2\pnucltr\pnstart1\pnindent720\pnhang {\pntxta .}}{\*\pnseclvl3\pndec\pnstart1\pnindent720\pnhang {\pntxta .}}{\*\pnseclvl4\pnlcltr\pnstart1\pnindent720\pnhang {\pntxta )}}{\*\pnseclvl5\pndec\pnstart1\pnindent720\pnhang {\pntxtb (}{\pntxta )}}
{\*\pnseclvl6\pnlcltr\pnstart1\pnindent720\pnhang {\pntxtb (}{\pntxta )}}{\*\pnseclvl7\pnlcrm\pnstart1\pnindent720\pnhang {\pntxtb (}{\pntxta )}}{\*\pnseclvl8\pnlcltr\pnstart1\pnindent720\pnhang {\pntxtb (}{\pntxta )}}{\*\pnseclvl9
\pnlcrm\pnstart1\pnindent720\pnhang {\pntxtb (}{\pntxta )}}\ltrrow\trowd \irow0\irowband0\ltrrow
\ts11\trgaph70\trrh300\trleft10\trftsWidth3\trwWidth10440\trautofit1\trpaddl70\trpaddr70\trpaddfl3\trpaddft3\trpaddfb3\trpaddfr3\tblrsid5536621\tbllkhdrrows\tbllkhdrcols\tbllknocolband\tblind80\tblindtype3 \clvertalc\clbrdrt\brdrs\brdrw20\brdrcf20 
\clbrdrl\brdrs\brdrw20\brdrcf20 \clbrdrb\brdrs\brdrw20\brdrcf20 \clbrdrr\brdrs\brdrw20\brdrcf20 \clcfpat1\clcbpat19\cltxlrtb\clftsWidth3\clwWidth10440\clcbpatraw19\clcfpatraw1\clhidemark \cellx10450\pard\plain \ltrpar
\qc \li0\ri0\widctlpar\intbl\wrapdefault\aspalpha\aspnum\faauto\adjustright\rin0\lin0\pararsid5536621 \rtlch\fcs1 \af0\afs24\alang1025 \ltrch\fcs0 \fs24\lang1033\langfe255\loch\af0\hich\af0\dbch\af31505\cgrid\langnp1033\langfenp255 {\rtlch\fcs1 
\af1\afs22 \ltrch\fcs0 \f1\fs22\cf8\lang3082\langfe3082\langnp3082\langfenp3082\insrsid5536621\charrsid5536621 Template para definir Rutinas (AOM)\cell }\pard \ltrpar\ql \li0\ri0\sa160\sl259\slmult1
\widctlpar\intbl\wrapdefault\aspalpha\aspnum\faauto\adjustright\rin0\lin0 {\rtlch\fcs1 \af0\afs20 \ltrch\fcs0 \fs20\lang3082\langfe3082\dbch\af0\langnp3082\langfenp3082\insrsid5536621\charrsid5536621 \trowd \irow0\irowband0\ltrrow
\ts11\trgaph70\trrh300\trleft10\trftsWidth3\trwWidth10440\trautofit1\trpaddl70\trpaddr70\trpaddfl3\trpaddft3\trpaddfb3\trpaddfr3\tblrsid5536621\tbllkhdrrows\tbllkhdrcols\tbllknocolband\tblind80\tblindtype3 \clvertalc\clbrdrt\brdrs\brdrw20\brdrcf20 
\clbrdrl\brdrs\brdrw20\brdrcf20 \clbrdrb\brdrs\brdrw20\brdrcf20 \clbrdrr\brdrs\brdrw20\brdrcf20 \clcfpat1\clcbpat19\cltxlrtb\clftsWidth3\clwWidth10440\clcbpatraw19\clcfpatraw1\clhidemark \cellx10450\row \ltrrow}\trowd \irow1\irowband1\ltrrow
\ts11\trgaph70\trrh540\trleft10\trftsWidth3\trwWidth10440\trautofit1\trpaddl70\trpaddr70\trpaddfl3\trpaddft3\trpaddfb3\trpaddfr3\tblrsid5536621\tbllkhdrrows\tbllkhdrcols\tbllknocolband\tblind80\tblindtype3 \clvertalc\clbrdrt\brdrnone \clbrdrl
\brdrs\brdrw20\brdrcf20 \clbrdrb\brdrs\brdrw20\brdrcf20 \clbrdrr\brdrs\brdrw20\brdrcf20 \clcfpat1\clcbpat19\cltxlrtb\clftsWidth3\clwWidth2060\clcbpatraw19\clcfpatraw1\clhidemark \cellx2070\clvertalc\clbrdrt\brdrnone \clbrdrl\brdrnone \clbrdrb
\brdrs\brdrw20\brdrcf20 \clbrdrr\brdrs\brdrw20\brdrcf20 \cltxlrtb\clftsWidth3\clwWidth8380\clhidemark \cellx10450\pard \ltrpar\ql \li0\ri0\widctlpar\intbl\wrapdefault\aspalpha\aspnum\faauto\adjustright\rin0\lin0\pararsid5536621 {\rtlch\fcs1 \af1\afs22 
\ltrch\fcs0 \f1\fs22\cf8\lang3082\langfe3082\langnp3082\langfenp3082\insrsid5536621\charrsid5536621 Contexto\cell }{\rtlch\fcs1 \af1\afs20 \ltrch\fcs0 \f1\fs20\cf1\lang3082\langfe3082\langnp3082\langfenp3082\insrsid5536621\charrsid5536621 \'bfDe }{
\rtlch\fcs1 \af1\afs20 \ltrch\fcs0 \f1\fs20\cf1\lang3082\langfe3082\langnp3082\langfenp3082\insrsid5536621\charrsid5536621 d\'f3nde}{\rtlch\fcs1 \af1\afs20 \ltrch\fcs0 
\f1\fs20\cf1\lang3082\langfe3082\langnp3082\langfenp3082\insrsid5536621\charrsid5536621  nace esta reuni\'f3n? \'bfComo se relaciona con las rutinas del }{\rtlch\fcs1 \af1\afs20 \ltrch\fcs0 
\f1\fs20\cf1\lang3082\langfe3082\langnp3082\langfenp3082\insrsid5536621\charrsid5536621 \'e1rea}{\rtlch\fcs1 \af1\afs20 \ltrch\fcs0 \f1\fs20\cf1\lang3082\langfe3082\langnp3082\langfenp3082\insrsid5536621\charrsid5536621 /gerencia/compa\'f1\'eda?\cell 
}\pard \ltrpar\ql \li0\ri0\sa160\sl259\slmult1\widctlpar\intbl\wrapdefault\aspalpha\aspnum\faauto\adjustright\rin0\lin0 {\rtlch\fcs1 \af0\afs20 \ltrch\fcs0 \fs20\lang3082\langfe3082\dbch\af0\langnp3082\langfenp3082\insrsid5536621\charrsid5536621 
\trowd \irow1\irowband1\ltrrow\ts11\trgaph70\trrh540\trleft10\trftsWidth3\trwWidth10440\trautofit1\trpaddl70\trpaddr70\trpaddfl3\trpaddft3\trpaddfb3\trpaddfr3\tblrsid5536621\tbllkhdrrows\tbllkhdrcols\tbllknocolband\tblind80\tblindtype3 \clvertalc\clbrdrt
\brdrnone \clbrdrl\brdrs\brdrw20\brdrcf20 \clbrdrb\brdrs\brdrw20\brdrcf20 \clbrdrr\brdrs\brdrw20\brdrcf20 \clcfpat1\clcbpat19\cltxlrtb\clftsWidth3\clwWidth2060\clcbpatraw19\clcfpatraw1\clhidemark \cellx2070\clvertalc\clbrdrt\brdrnone \clbrdrl\brdrnone 
\clbrdrb\brdrs\brdrw20\brdrcf20 \clbrdrr\brdrs\brdrw20\brdrcf20 \cltxlrtb\clftsWidth3\clwWidth8380\clhidemark \cellx10450\row \ltrrow}\trowd \irow2\irowband2\ltrrow
\ts11\trgaph70\trrh300\trleft10\trftsWidth3\trwWidth10440\trautofit1\trpaddl70\trpaddr70\trpaddfl3\trpaddft3\trpaddfb3\trpaddfr3\tblrsid5536621\tbllkhdrrows\tbllkhdrcols\tbllknocolband\tblind80\tblindtype3 \clvertalc\clbrdrt\brdrnone \clbrdrl
\brdrs\brdrw20\brdrcf20 \clbrdrb\brdrs\brdrw20\brdrcf20 \clbrdrr\brdrs\brdrw20\brdrcf20 \clcfpat1\clcbpat19\cltxlrtb\clftsWidth3\clwWidth2060\clcbpatraw19\clcfpatraw1\clhidemark \cellx2070\clvertalc\clbrdrt\brdrnone \clbrdrl\brdrnone \clbrdrb
\brdrs\brdrw20\brdrcf20 \clbrdrr\brdrs\brdrw20\brdrcf20 \cltxlrtb\clftsWidth3\clwWidth8380\clhidemark \cellx10450\pard \ltrpar\ql \li0\ri0\widctlpar\intbl\wrapdefault\aspalpha\aspnum\faauto\adjustright\rin0\lin0\pararsid5536621 {\rtlch\fcs1 \af1\afs22 
\ltrch\fcs0 \f1\fs22\cf8\lang3082\langfe3082\langnp3082\langfenp3082\insrsid5536621\charrsid5536621 Prop\'f3sito\cell }{\rtlch\fcs1 \af1\afs20 \ltrch\fcs0 \f1\fs20\cf1\lang3082\langfe3082\langnp3082\langfenp3082\insrsid5536621\charrsid5536621 \'bfPor qu
\'e9 existe esta reuni\'f3n? \'bfCu\'e1l es el }{\rtlch\fcs1 \af1\afs20 \ltrch\fcs0 \f1\fs20\cf1\lang3082\langfe3082\langnp3082\langfenp3082\insrsid5536621\charrsid5536621 prop\'f3sito}{\rtlch\fcs1 \af1\afs20 \ltrch\fcs0 
\f1\fs20\cf1\lang3082\langfe3082\langnp3082\langfenp3082\insrsid5536621\charrsid5536621 ?\cell }\pard \ltrpar\ql \li0\ri0\sa160\sl259\slmult1\widctlpar\intbl\wrapdefault\aspalpha\aspnum\faauto\adjustright\rin0\lin0 {\rtlch\fcs1 \af0\afs20 \ltrch\fcs0 
\fs20\lang3082\langfe3082\dbch\af0\langnp3082\langfenp3082\insrsid5536621\charrsid5536621 \trowd \irow2\irowband2\ltrrow
\ts11\trgaph70\trrh300\trleft10\trftsWidth3\trwWidth10440\trautofit1\trpaddl70\trpaddr70\trpaddfl3\trpaddft3\trpaddfb3\trpaddfr3\tblrsid5536621\tbllkhdrrows\tbllkhdrcols\tbllknocolband\tblind80\tblindtype3 \clvertalc\clbrdrt\brdrnone \clbrdrl
\brdrs\brdrw20\brdrcf20 \clbrdrb\brdrs\brdrw20\brdrcf20 \clbrdrr\brdrs\brdrw20\brdrcf20 \clcfpat1\clcbpat19\cltxlrtb\clftsWidth3\clwWidth2060\clcbpatraw19\clcfpatraw1\clhidemark \cellx2070\clvertalc\clbrdrt\brdrnone \clbrdrl\brdrnone \clbrdrb
\brdrs\brdrw20\brdrcf20 \clbrdrr\brdrs\brdrw20\brdrcf20 \cltxlrtb\clftsWidth3\clwWidth8380\clhidemark \cellx10450\row \ltrrow}\trowd \irow3\irowband3\ltrrow
\ts11\trgaph70\trrh540\trleft10\trftsWidth3\trwWidth10440\trautofit1\trpaddl70\trpaddr70\trpaddfl3\trpaddft3\trpaddfb3\trpaddfr3\tblrsid5536621\tbllkhdrrows\tbllkhdrcols\tbllknocolband\tblind80\tblindtype3 \clvertalc\clbrdrt\brdrnone \clbrdrl
\brdrs\brdrw20\brdrcf20 \clbrdrb\brdrs\brdrw20\brdrcf20 \clbrdrr\brdrs\brdrw20\brdrcf20 \clcfpat1\clcbpat19\cltxlrtb\clftsWidth3\clwWidth2060\clcbpatraw19\clcfpatraw1\clhidemark \cellx2070\clvertalc\clbrdrt\brdrnone \clbrdrl\brdrnone \clbrdrb
\brdrs\brdrw20\brdrcf20 \clbrdrr\brdrs\brdrw20\brdrcf20 \cltxlrtb\clftsWidth3\clwWidth8380\clhidemark \cellx10450\pard \ltrpar\ql \li0\ri0\widctlpar\intbl\wrapdefault\aspalpha\aspnum\faauto\adjustright\rin0\lin0\pararsid5536621 {\rtlch\fcs1 \af1\afs22 
\ltrch\fcs0 \f1\fs22\cf8\lang3082\langfe3082\langnp3082\langfenp3082\insrsid5536621\charrsid5536621 Cantidad\cell }{\rtlch\fcs1 \af1\afs20 \ltrch\fcs0 \f1\fs20\cf1\lang3082\langfe3082\langnp3082\langfenp3082\insrsid5536621\charrsid5536621 
Por ejemplo: dos reuniones por semana para validar acciones semanales.\cell }\pard \ltrpar\ql \li0\ri0\sa160\sl259\slmult1\widctlpar\intbl\wrapdefault\aspalpha\aspnum\faauto\adjustright\rin0\lin0 {\rtlch\fcs1 \af0\afs20 \ltrch\fcs0 
\fs20\lang3082\langfe3082\dbch\af0\langnp3082\langfenp3082\insrsid5536621\charrsid5536621 \trowd \irow3\irowband3\ltrrow
\ts11\trgaph70\trrh540\trleft10\trftsWidth3\trwWidth10440\trautofit1\trpaddl70\trpaddr70\trpaddfl3\trpaddft3\trpaddfb3\trpaddfr3\tblrsid5536621\tbllkhdrrows\tbllkhdrcols\tbllknocolband\tblind80\tblindtype3 \clvertalc\clbrdrt\brdrnone \clbrdrl
\brdrs\brdrw20\brdrcf20 \clbrdrb\brdrs\brdrw20\brdrcf20 \clbrdrr\brdrs\brdrw20\brdrcf20 \clcfpat1\clcbpat19\cltxlrtb\clftsWidth3\clwWidth2060\clcbpatraw19\clcfpatraw1\clhidemark \cellx2070\clvertalc\clbrdrt\brdrnone \clbrdrl\brdrnone \clbrdrb
\brdrs\brdrw20\brdrcf20 \clbrdrr\brdrs\brdrw20\brdrcf20 \cltxlrtb\clftsWidth3\clwWidth8380\clhidemark \cellx10450\row \ltrrow}\trowd \irow4\irowband4\ltrrow
\ts11\trgaph70\trrh300\trleft10\trftsWidth3\trwWidth10440\trautofit1\trpaddl70\trpaddr70\trpaddfl3\trpaddft3\trpaddfb3\trpaddfr3\tblrsid5536621\tbllkhdrrows\tbllkhdrcols\tbllknocolband\tblind80\tblindtype3 \clvertalc\clbrdrt\brdrnone \clbrdrl
\brdrs\brdrw20\brdrcf20 \clbrdrb\brdrs\brdrw20\brdrcf20 \clbrdrr\brdrs\brdrw20\brdrcf20 \clcfpat1\clcbpat19\cltxlrtb\clftsWidth3\clwWidth2060\clcbpatraw19\clcfpatraw1\clhidemark \cellx2070\clvertalc\clbrdrt\brdrnone \clbrdrl\brdrnone \clbrdrb
\brdrs\brdrw20\brdrcf20 \clbrdrr\brdrs\brdrw20\brdrcf20 \cltxlrtb\clftsWidth3\clwWidth8380\clhidemark \cellx10450\pard \ltrpar\ql \li0\ri0\widctlpar\intbl\wrapdefault\aspalpha\aspnum\faauto\adjustright\rin0\lin0\pararsid5536621 {\rtlch\fcs1 \af1\afs22 
\ltrch\fcs0 \f1\fs22\cf8\lang3082\langfe3082\langnp3082\langfenp3082\insrsid5536621\charrsid5536621 Calidad\cell }{\rtlch\fcs1 \af1\afs20 \ltrch\fcs0 \f1\fs20\cf1\lang3082\langfe3082\langnp3082\langfenp3082\insrsid5536621\charrsid5536621 \'bfQu\'e9
 se necesita para poder lograr los }{\rtlch\fcs1 \af1\afs20 \ltrch\fcs0 \f1\fs20\cf1\lang3082\langfe3082\langnp3082\langfenp3082\insrsid5536621\charrsid5536621 objetivos}{\rtlch\fcs1 \af1\afs20 \ltrch\fcs0 
\f1\fs20\cf1\lang3082\langfe3082\langnp3082\langfenp3082\insrsid5536621\charrsid5536621 ? \cell }\pard \ltrpar\ql \li0\ri0\sa160\sl259\slmult1\widctlpar\intbl\wrapdefault\aspalpha\aspnum\faauto\adjustright\rin0\lin0 {\rtlch\fcs1 \af0\afs20 \ltrch\fcs0 
\fs20\lang3082\langfe3082\dbch\af0\langnp3082\langfenp3082\insrsid5536621\charrsid5536621 \trowd \irow4\irowband4\ltrrow
\ts11\trgaph70\trrh300\trleft10\trftsWidth3\trwWidth10440\trautofit1\trpaddl70\trpaddr70\trpaddfl3\trpaddft3\trpaddfb3\trpaddfr3\tblrsid5536621\tbllkhdrrows\tbllkhdrcols\tbllknocolband\tblind80\tblindtype3 \clvertalc\clbrdrt\brdrnone \clbrdrl
\brdrs\brdrw20\brdrcf20 \clbrdrb\brdrs\brdrw20\brdrcf20 \clbrdrr\brdrs\brdrw20\brdrcf20 \clcfpat1\clcbpat19\cltxlrtb\clftsWidth3\clwWidth2060\clcbpatraw19\clcfpatraw1\clhidemark \cellx2070\clvertalc\clbrdrt\brdrnone \clbrdrl\brdrnone \clbrdrb
\brdrs\brdrw20\brdrcf20 \clbrdrr\brdrs\brdrw20\brdrcf20 \cltxlrtb\clftsWidth3\clwWidth8380\clhidemark \cellx10450\row \ltrrow}\trowd \irow5\irowband5\ltrrow
\ts11\trgaph70\trrh300\trleft10\trftsWidth3\trwWidth10440\trautofit1\trpaddl70\trpaddr70\trpaddfl3\trpaddft3\trpaddfb3\trpaddfr3\tblrsid5536621\tbllkhdrrows\tbllkhdrcols\tbllknocolband\tblind80\tblindtype3 \clvertalc\clbrdrt\brdrnone \clbrdrl
\brdrs\brdrw20\brdrcf20 \clbrdrb\brdrs\brdrw20\brdrcf20 \clbrdrr\brdrs\brdrw20\brdrcf20 \clcfpat1\clcbpat19\cltxlrtb\clftsWidth3\clwWidth2060\clcbpatraw19\clcfpatraw1\clhidemark \cellx2070\clvertalc\clbrdrt\brdrnone \clbrdrl\brdrnone \clbrdrb\brdrnone 
\clbrdrr\brdrs\brdrw20\brdrcf20 \cltxlrtb\clftsWidth3\clwWidth8380\clhidemark \cellx10450\pard \ltrpar\ql \li0\ri0\widctlpar\intbl\wrapdefault\aspalpha\aspnum\faauto\adjustright\rin0\lin0\pararsid5536621 {\rtlch\fcs1 \af1\afs22 \ltrch\fcs0 
\f1\fs22\cf8\lang3082\langfe3082\langnp3082\langfenp3082\insrsid5536621\charrsid5536621 Recursos\cell }{\rtlch\fcs1 \ai\af1\afs18 \ltrch\fcs0 \i\f1\fs18\cf1\lang3082\langfe3082\langnp3082\langfenp3082\insrsid5536621\charrsid5536621 \'bfQu\'e9
 personas son necesarias para esta reuni\'f3n?\cell }\pard \ltrpar\ql \li0\ri0\sa160\sl259\slmult1\widctlpar\intbl\wrapdefault\aspalpha\aspnum\faauto\adjustright\rin0\lin0 {\rtlch\fcs1 \af0\afs20 \ltrch\fcs0 
\fs20\lang3082\langfe3082\dbch\af0\langnp3082\langfenp3082\insrsid5536621\charrsid5536621 \trowd \irow5\irowband5\ltrrow
\ts11\trgaph70\trrh300\trleft10\trftsWidth3\trwWidth10440\trautofit1\trpaddl70\trpaddr70\trpaddfl3\trpaddft3\trpaddfb3\trpaddfr3\tblrsid5536621\tbllkhdrrows\tbllkhdrcols\tbllknocolband\tblind80\tblindtype3 \clvertalc\clbrdrt\brdrnone \clbrdrl
\brdrs\brdrw20\brdrcf20 \clbrdrb\brdrs\brdrw20\brdrcf20 \clbrdrr\brdrs\brdrw20\brdrcf20 \clcfpat1\clcbpat19\cltxlrtb\clftsWidth3\clwWidth2060\clcbpatraw19\clcfpatraw1\clhidemark \cellx2070\clvertalc\clbrdrt\brdrnone \clbrdrl\brdrnone \clbrdrb\brdrnone 
\clbrdrr\brdrs\brdrw20\brdrcf20 \cltxlrtb\clftsWidth3\clwWidth8380\clhidemark \cellx10450\row \ltrrow}\trowd \irow6\irowband6\ltrrow
\ts11\trgaph70\trrh300\trleft10\trftsWidth3\trwWidth10440\trautofit1\trpaddl70\trpaddr70\trpaddfl3\trpaddft3\trpaddfb3\trpaddfr3\tblrsid5536621\tbllkhdrrows\tbllkhdrcols\tbllknocolband\tblind80\tblindtype3 \clvertalc\clbrdrt\brdrnone \clbrdrl
\brdrs\brdrw20\brdrcf20 \clbrdrb\brdrs\brdrw20\brdrcf20 \clbrdrr\brdrs\brdrw20\brdrcf20 \clcfpat1\clcbpat19\cltxlrtb\clftsWidth3\clwWidth2060\clcbpatraw19\clcfpatraw1\clhidemark \cellx2070\clvertalc\clbrdrt\brdrs\brdrw20\brdrcf20 \clbrdrl\brdrnone 
\clbrdrb\brdrs\brdrw20\brdrcf20 \clbrdrr\brdrs\brdrw20\brdrcf20 \cltxlrtb\clftsWidth3\clwWidth8380\clhidemark \cellx10450\pard \ltrpar\ql \li0\ri0\widctlpar\intbl\wrapdefault\aspalpha\aspnum\faauto\adjustright\rin0\lin0\pararsid5536621 {\rtlch\fcs1 
\af1\afs22 \ltrch\fcs0 \f1\fs22\cf8\lang3082\langfe3082\langnp3082\langfenp3082\insrsid5536621\charrsid5536621 Tiempo\cell }{\rtlch\fcs1 \ai\af1\afs18 \ltrch\fcs0 \i\f1\fs18\cf15\lang3082\langfe3082\langnp3082\langfenp3082\insrsid5536621\charrsid5536621 
Horario\cell }\pard \ltrpar\ql \li0\ri0\sa160\sl259\slmult1\widctlpar\intbl\wrapdefault\aspalpha\aspnum\faauto\adjustright\rin0\lin0 {\rtlch\fcs1 \af0\afs20 \ltrch\fcs0 
\fs20\lang3082\langfe3082\dbch\af0\langnp3082\langfenp3082\insrsid5536621\charrsid5536621 \trowd \irow6\irowband6\ltrrow
\ts11\trgaph70\trrh300\trleft10\trftsWidth3\trwWidth10440\trautofit1\trpaddl70\trpaddr70\trpaddfl3\trpaddft3\trpaddfb3\trpaddfr3\tblrsid5536621\tbllkhdrrows\tbllkhdrcols\tbllknocolband\tblind80\tblindtype3 \clvertalc\clbrdrt\brdrnone \clbrdrl
\brdrs\brdrw20\brdrcf20 \clbrdrb\brdrs\brdrw20\brdrcf20 \clbrdrr\brdrs\brdrw20\brdrcf20 \clcfpat1\clcbpat19\cltxlrtb\clftsWidth3\clwWidth2060\clcbpatraw19\clcfpatraw1\clhidemark \cellx2070\clvertalc\clbrdrt\brdrs\brdrw20\brdrcf20 \clbrdrl\brdrnone 
\clbrdrb\brdrs\brdrw20\brdrcf20 \clbrdrr\brdrs\brdrw20\brdrcf20 \cltxlrtb\clftsWidth3\clwWidth8380\clhidemark \cellx10450\row \ltrrow}\trowd \irow7\irowband7\ltrrow
\ts11\trgaph70\trrh288\trleft10\trftsWidth3\trwWidth10440\trautofit1\trpaddl70\trpaddr70\trpaddfl3\trpaddft3\trpaddfb3\trpaddfr3\tblrsid5536621\tbllkhdrrows\tbllkhdrcols\tbllknocolband\tblind80\tblindtype3 \clvmgf\clvertalc\clbrdrt\brdrnone \clbrdrl
\brdrs\brdrw20\brdrcf20 \clbrdrb\brdrs\brdrw20\brdrcf20 \clbrdrr\brdrs\brdrw20\brdrcf20 \clcfpat1\clcbpat19\cltxlrtb\clftsWidth3\clwWidth2060\clcbpatraw19\clcfpatraw1\clhidemark \cellx2070\clvertalc\clbrdrt\brdrnone \clbrdrl\brdrnone \clbrdrb\brdrnone 
\clbrdrr\brdrs\brdrw20\brdrcf20 \cltxlrtb\clftsWidth3\clwWidth8380\clhidemark \cellx10450\pard \ltrpar\ql \li0\ri0\widctlpar\intbl\wrapdefault\aspalpha\aspnum\faauto\adjustright\rin0\lin0\pararsid5536621 {\rtlch\fcs1 \af1\afs22 \ltrch\fcs0 
\f1\fs22\cf8\lang3082\langfe3082\langnp3082\langfenp3082\insrsid5536621\charrsid5536621 Agenda\cell }{\rtlch\fcs1 \ai\af1\afs18 \ltrch\fcs0 \i\f1\fs18\cf1\lang3082\langfe3082\langnp3082\langfenp3082\insrsid5536621\charrsid5536621 
Tema 1 (Tiempo 1) - Responsable 1\cell }\pard \ltrpar\ql \li0\ri0\sa160\sl259\slmult1\widctlpar\intbl\wrapdefault\aspalpha\aspnum\faauto\adjustright\rin0\lin0 {\rtlch\fcs1 \af0\afs20 \ltrch\fcs0 
\fs20\lang3082\langfe3082\dbch\af0\langnp3082\langfenp3082\insrsid5536621\charrsid5536621 \trowd \irow7\irowband7\ltrrow
\ts11\trgaph70\trrh288\trleft10\trftsWidth3\trwWidth10440\trautofit1\trpaddl70\trpaddr70\trpaddfl3\trpaddft3\trpaddfb3\trpaddfr3\tblrsid5536621\tbllkhdrrows\tbllkhdrcols\tbllknocolband\tblind80\tblindtype3 \clvmgf\clvertalc\clbrdrt\brdrnone \clbrdrl
\brdrs\brdrw20\brdrcf20 \clbrdrb\brdrs\brdrw20\brdrcf20 \clbrdrr\brdrs\brdrw20\brdrcf20 \clcfpat1\clcbpat19\cltxlrtb\clftsWidth3\clwWidth2060\clcbpatraw19\clcfpatraw1\clhidemark \cellx2070\clvertalc\clbrdrt\brdrnone \clbrdrl\brdrnone \clbrdrb\brdrnone 
\clbrdrr\brdrs\brdrw20\brdrcf20 \cltxlrtb\clftsWidth3\clwWidth8380\clhidemark \cellx10450\row \ltrrow}\trowd \irow8\irowband8\ltrrow
\ts11\trgaph70\trrh288\trleft10\trftsWidth3\trwWidth10440\trautofit1\trpaddl70\trpaddr70\trpaddfl3\trpaddft3\trpaddfb3\trpaddfr3\tblrsid5536621\tbllkhdrrows\tbllkhdrcols\tbllknocolband\tblind80\tblindtype3 \clvmrg\clvertalc\clbrdrt\brdrnone \clbrdrl
\brdrs\brdrw20\brdrcf20 \clbrdrb\brdrs\brdrw20\brdrcf20 \clbrdrr\brdrs\brdrw20\brdrcf20 \cltxlrtb\clftsWidth3\clwWidth2060\clshdrawnil\clhidemark \cellx2070\clvertalc\clbrdrt\brdrnone \clbrdrl\brdrnone \clbrdrb\brdrnone \clbrdrr\brdrs\brdrw20\brdrcf20 
\cltxlrtb\clftsWidth3\clwWidth8380\clhidemark \cellx10450\pard \ltrpar\ql \li0\ri0\widctlpar\intbl\wrapdefault\aspalpha\aspnum\faauto\adjustright\rin0\lin0\pararsid5536621 {\rtlch\fcs1 \af1\afs22 \ltrch\fcs0 
\f1\fs22\cf8\lang3082\langfe3082\langnp3082\langfenp3082\insrsid5536621\charrsid5536621 \cell }{\rtlch\fcs1 \ai\af1\afs18 \ltrch\fcs0 \i\f1\fs18\cf1\lang3082\langfe3082\langnp3082\langfenp3082\insrsid5536621\charrsid5536621 Tema 2 (Tiempo 2) - Respo
nsable 2\cell }\pard \ltrpar\ql \li0\ri0\sa160\sl259\slmult1\widctlpar\intbl\wrapdefault\aspalpha\aspnum\faauto\adjustright\rin0\lin0 {\rtlch\fcs1 \af0\afs20 \ltrch\fcs0 
\fs20\lang3082\langfe3082\dbch\af0\langnp3082\langfenp3082\insrsid5536621\charrsid5536621 \trowd \irow8\irowband8\ltrrow
\ts11\trgaph70\trrh288\trleft10\trftsWidth3\trwWidth10440\trautofit1\trpaddl70\trpaddr70\trpaddfl3\trpaddft3\trpaddfb3\trpaddfr3\tblrsid5536621\tbllkhdrrows\tbllkhdrcols\tbllknocolband\tblind80\tblindtype3 \clvmrg\clvertalc\clbrdrt\brdrnone \clbrdrl
\brdrs\brdrw20\brdrcf20 \clbrdrb\brdrs\brdrw20\brdrcf20 \clbrdrr\brdrs\brdrw20\brdrcf20 \cltxlrtb\clftsWidth3\clwWidth2060\clshdrawnil\clhidemark \cellx2070\clvertalc\clbrdrt\brdrnone \clbrdrl\brdrnone \clbrdrb\brdrnone \clbrdrr\brdrs\brdrw20\brdrcf20 
\cltxlrtb\clftsWidth3\clwWidth8380\clhidemark \cellx10450\row \ltrrow}\pard \ltrpar\ql \li0\ri0\widctlpar\intbl\wrapdefault\aspalpha\aspnum\faauto\adjustright\rin0\lin0\pararsid5536621 {\rtlch\fcs1 \af1\afs22 \ltrch\fcs0 
\f1\fs22\cf8\lang3082\langfe3082\langnp3082\langfenp3082\insrsid5536621\charrsid5536621 \cell }{\rtlch\fcs1 \ai\af1\afs18 \ltrch\fcs0 \i\f1\fs18\cf1\lang3082\langfe3082\langnp3082\langfenp3082\insrsid5536621\charrsid5536621 
Tema 3 (Tiempo 3) - Responsable 3\cell }\pard \ltrpar\ql \li0\ri0\sa160\sl259\slmult1\widctlpar\intbl\wrapdefault\aspalpha\aspnum\faauto\adjustright\rin0\lin0 {\rtlch\fcs1 \af0\afs20 \ltrch\fcs0 
\fs20\lang3082\langfe3082\dbch\af0\langnp3082\langfenp3082\insrsid5536621\charrsid5536621 \trowd \irow9\irowband9\ltrrow
\ts11\trgaph70\trrh288\trleft10\trftsWidth3\trwWidth10440\trautofit1\trpaddl70\trpaddr70\trpaddfl3\trpaddft3\trpaddfb3\trpaddfr3\tblrsid5536621\tbllkhdrrows\tbllkhdrcols\tbllknocolband\tblind80\tblindtype3 \clvmrg\clvertalc\clbrdrt\brdrnone \clbrdrl
\brdrs\brdrw20\brdrcf20 \clbrdrb\brdrs\brdrw20\brdrcf20 \clbrdrr\brdrs\brdrw20\brdrcf20 \cltxlrtb\clftsWidth3\clwWidth2060\clshdrawnil\clhidemark \cellx2070\clvertalc\clbrdrt\brdrnone \clbrdrl\brdrnone \clbrdrb\brdrnone \clbrdrr\brdrs\brdrw20\brdrcf20 
\cltxlrtb\clftsWidth3\clwWidth8380\clhidemark \cellx10450\row \ltrrow}\trowd \irow10\irowband10\lastrow \ltrrow
\ts11\trgaph70\trrh300\trleft10\trftsWidth3\trwWidth10440\trautofit1\trpaddl70\trpaddr70\trpaddfl3\trpaddft3\trpaddfb3\trpaddfr3\tblrsid5536621\tbllkhdrrows\tbllkhdrcols\tbllknocolband\tblind80\tblindtype3 \clvmrg\clvertalc\clbrdrt\brdrnone \clbrdrl
\brdrs\brdrw20\brdrcf20 \clbrdrb\brdrs\brdrw20\brdrcf20 \clbrdrr\brdrs\brdrw20\brdrcf20 \cltxlrtb\clftsWidth3\clwWidth2060\clshdrawnil\clhidemark \cellx2070\clvertalc\clbrdrt\brdrnone \clbrdrl\brdrnone \clbrdrb\brdrs\brdrw20\brdrcf20 \clbrdrr
\brdrs\brdrw20\brdrcf20 \cltxlrtb\clftsWidth3\clwWidth8380\clhidemark \cellx10450\pard \ltrpar\ql \li0\ri0\widctlpar\intbl\wrapdefault\aspalpha\aspnum\faauto\adjustright\rin0\lin0\pararsid5536621 {\rtlch\fcs1 \af1\afs22 \ltrch\fcs0 
\f1\fs22\cf8\lang3082\langfe3082\langnp3082\langfenp3082\insrsid5536621\charrsid5536621 \cell }{\rtlch\fcs1 \ai\af1\afs18 \ltrch\fcs0 \i\f1\fs18\cf15\lang3082\langfe3082\langnp3082\langfenp3082\insrsid5536621\charrsid5536621 \~\cell }\pard \ltrpar
\ql \li0\ri0\sa160\sl259\slmult1\widctlpar\intbl\wrapdefault\aspalpha\aspnum\faauto\adjustright\rin0\lin0 {\rtlch\fcs1 \af0\afs20 \ltrch\fcs0 \fs20\lang3082\langfe3082\dbch\af0\langnp3082\langfenp3082\insrsid5536621\charrsid5536621 
\trowd \irow10\irowband10\lastrow \ltrrow\ts11\trgaph70\trrh300\trleft10\trftsWidth3\trwWidth10440\trautofit1\trpaddl70\trpaddr70\trpaddfl3\trpaddft3\trpaddfb3\trpaddfr3\tblrsid5536621\tbllkhdrrows\tbllkhdrcols\tbllknocolband\tblind80\tblindtype3 
\clvmrg\clvertalc\clbrdrt\brdrnone \clbrdrl\brdrs\brdrw20\brdrcf20 \clbrdrb\brdrs\brdrw20\brdrcf20 \clbrdrr\brdrs\brdrw20\brdrcf20 \cltxlrtb\clftsWidth3\clwWidth2060\clshdrawnil\clhidemark \cellx2070\clvertalc\clbrdrt\brdrnone \clbrdrl\brdrnone \clbrdrb
\brdrs\brdrw20\brdrcf20 \clbrdrr\brdrs\brdrw20\brdrcf20 \cltxlrtb\clftsWidth3\clwWidth8380\clhidemark \cellx10450\row }\pard \ltrpar\ql \li0\ri0\nowidctlpar\wrapdefault\faauto\rin0\lin0\itap0\pararsid12210631 {\rtlch\fcs1 \af0 \ltrch\fcs0 
\lang3082\langfe255\langnp3082\insrsid15217017\charrsid14426457 
\par }";

                    newCita.RTFBody = System.Text.Encoding.UTF8.GetBytes(rtb.Rtf);

                    newCita.Start = DateTime.Now.AddHours(2);
                    newCita.End = DateTime.Now.AddHours(3);
                    newCita.Location = "Elija la ubicación de la Reunión";
                    newCita.Subject = "Reunión Template";
                    newCita.Recipients.Add("Danilo Delgado Redlich");
                    Microsoft.Office.Interop.Outlook.Recipients sentTo = newCita.Recipients;
                    Microsoft.Office.Interop.Outlook.Recipient sentInvite = null;
                    sentInvite = sentTo.Add("Juanito Perez");
                    sentInvite.Type = (int)Microsoft.Office.Interop.Outlook.OlMeetingRecipientType.olRequired;
                    sentInvite = sentTo.Add("Luis Mario Gonzalez");
                    sentInvite.Type = (int)Microsoft.Office.Interop.Outlook.OlMeetingRecipientType.olOptional;
                    sentTo.ResolveAll();
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
                    string HTMLTemplate = "<table style='border-collapse:collapse;border-spacing:0;table-layout: fixed; width: 1160px' class='tg'><colgroup><col style='width: 27px'><col style='width: 390px'><col style='width: 244px'><col style='width: 193px'><col style='width: 306px'></colgroup><thead><tr><th style='background-color:#0075b0;border-color:#000000;border-style:solid;border-width:1px;color:#ffffff;font-family:Arial, sans-serif;font-size:14px;font-weight:normal;overflow:hidden;padding:10px 5px;text-align:center;vertical-align:top;word-break:normal' colspan='5'>Guía Template para Minutas de Acciones</th></tr></thead><tbody><tr><td style='background-color:#002776;border-color:#000000;border-style:solid;border-width:1px;color:#ffffff;font-family:Arial, sans-serif;font-size:14px;overflow:hidden;padding:10px 5px;text-align:left;vertical-align:top;word-break:normal'>N°</td><td style='background-color:#002776;border-color:#000000;border-style:solid;border-width:1px;color:#ffffff;font-family:Arial, sans-serif;font-size:14px;overflow:hidden;padding:10px 5px;text-align:left;vertical-align:top;word-break:normal'>Tema y Descripción de la Acción</td><td style='background-color:#002776;border-color:#000000;border-style:solid;border-width:1px;color:#ffffff;font-family:Arial, sans-serif;font-size:14px;overflow:hidden;padding:10px 5px;text-align:left;vertical-align:top;word-break:normal'>Persona Responsable</td><td style='background-color:#002776;border-color:#000000;border-style:solid;border-width:1px;color:#ffffff;font-family:Arial, sans-serif;font-size:14px;overflow:hidden;padding:10px 5px;text-align:left;vertical-align:top;word-break:normal'>Fecha Límite</td><td style='background-color:#002776;border-color:#000000;border-style:solid;border-width:1px;color:#ffffff;font-family:Arial, sans-serif;font-size:14px;overflow:hidden;padding:10px 5px;text-align:left;vertical-align:top;word-break:normal'>Status</td></tr><tr><td style='background-color:#002776;border-color:#000000;border-style:solid;border-width:1px;color:#ffffff;font-family:Arial, sans-serif;font-size:14px;overflow:hidden;padding:10px 5px;text-align:left;vertical-align:top;word-break:normal'>1</td><td style='border-color:#000000;border-style:solid;border-width:1px;font-family:Arial, sans-serif;font-size:14px;overflow:hidden;padding:10px 5px;text-align:left;vertical-align:top;word-break:normal'></td><td style='border-color:#000000;border-style:solid;border-width:1px;font-family:Arial, sans-serif;font-size:14px;overflow:hidden;padding:10px 5px;text-align:left;vertical-align:top;word-break:normal'></td><td style='border-color:#000000;border-style:solid;border-width:1px;font-family:Arial, sans-serif;font-size:14px;overflow:hidden;padding:10px 5px;text-align:left;vertical-align:top;word-break:normal'></td><td style='border-color:#000000;border-style:solid;border-width:1px;font-family:Arial, sans-serif;font-size:14px;overflow:hidden;padding:10px 5px;text-align:left;vertical-align:top;word-break:normal'></td></tr><tr><td style='background-color:#002776;border-color:#000000;border-style:solid;border-width:1px;color:#ffffff;font-family:Arial, sans-serif;font-size:14px;overflow:hidden;padding:10px 5px;text-align:left;vertical-align:top;word-break:normal'>2</td><td style='border-color:#000000;border-style:solid;border-width:1px;font-family:Arial, sans-serif;font-size:14px;overflow:hidden;padding:10px 5px;text-align:left;vertical-align:top;word-break:normal'></td><td style='border-color:#000000;border-style:solid;border-width:1px;font-family:Arial, sans-serif;font-size:14px;overflow:hidden;padding:10px 5px;text-align:left;vertical-align:top;word-break:normal'></td><td style='border-color:#000000;border-style:solid;border-width:1px;font-family:Arial, sans-serif;font-size:14px;overflow:hidden;padding:10px 5px;text-align:left;vertical-align:top;word-break:normal'></td><td style='border-color:#000000;border-style:solid;border-width:1px;font-family:Arial, sans-serif;font-size:14px;overflow:hidden;padding:10px 5px;text-align:left;vertical-align:top;word-break:normal'></td></tr><tr><td style='background-color:#002776;border-color:#000000;border-style:solid;border-width:1px;color:#ffffff;font-family:Arial, sans-serif;font-size:14px;overflow:hidden;padding:10px 5px;text-align:left;vertical-align:top;word-break:normal'>3</td><td style='border-color:#000000;border-style:solid;border-width:1px;font-family:Arial, sans-serif;font-size:14px;overflow:hidden;padding:10px 5px;text-align:left;vertical-align:top;word-break:normal'></td><td style='border-color:#000000;border-style:solid;border-width:1px;font-family:Arial, sans-serif;font-size:14px;overflow:hidden;padding:10px 5px;text-align:left;vertical-align:top;word-break:normal'></td><td style='border-color:#000000;border-style:solid;border-width:1px;font-family:Arial, sans-serif;font-size:14px;overflow:hidden;padding:10px 5px;text-align:left;vertical-align:top;word-break:normal'></td><td style='border-color:#000000;border-style:solid;border-width:1px;font-family:Arial, sans-serif;font-size:14px;overflow:hidden;padding:10px 5px;text-align:left;vertical-align:top;word-break:normal'></td></tr><tr><td style='background-color:#002776;border-color:#000000;border-style:solid;border-width:1px;color:#ffffff;font-family:Arial, sans-serif;font-size:14px;overflow:hidden;padding:10px 5px;text-align:left;vertical-align:top;word-break:normal'>4</td><td style='border-color:#000000;border-style:solid;border-width:1px;font-family:Arial, sans-serif;font-size:14px;overflow:hidden;padding:10px 5px;text-align:left;vertical-align:top;word-break:normal'></td><td style='border-color:#000000;border-style:solid;border-width:1px;font-family:Arial, sans-serif;font-size:14px;overflow:hidden;padding:10px 5px;text-align:left;vertical-align:top;word-break:normal'></td><td style='border-color:#000000;border-style:solid;border-width:1px;font-family:Arial, sans-serif;font-size:14px;overflow:hidden;padding:10px 5px;text-align:left;vertical-align:top;word-break:normal'></td><td style='border-color:#000000;border-style:solid;border-width:1px;font-family:Arial, sans-serif;font-size:14px;overflow:hidden;padding:10px 5px;text-align:left;vertical-align:top;word-break:normal'></td></tr></tbody></table>";
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
