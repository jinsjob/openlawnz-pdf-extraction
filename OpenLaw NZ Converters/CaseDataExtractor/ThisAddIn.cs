using Shared;
using System;
using Word = Microsoft.Office.Interop.Word;

namespace CaseDataExtractor
{


    public partial class ThisAddIn
    {

        private void process(Word.Document Doc)
        {
            if (Doc.FullName.Contains("__openlawnz_from_pdf"))
            {
                this.Application.ScreenUpdating = false;
                this.Application.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;

                string logPath = Doc.FullName.Replace(".docx", ".txt");
                Logger logger = new Logger(logPath);
                try
                {
                    ExtractFootnotes.Process(Doc, this.Application, logger);
                }
                catch
                {

                }
                this.Application.ScreenUpdating = true;
                this.Application.Quit(false);
            }
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {

            Word.ApplicationEvents2_Event wdEvents2 = (Word.ApplicationEvents2_Event)this.Application;
            wdEvents2.DocumentOpen += WdEvents2_DocumentOpen;

            try
            {
                Word.Document Doc = this.Application.ActiveDocument;
                if (String.IsNullOrWhiteSpace(Doc.Path))
                {
                    Console.WriteLine(String.Format("Word initialized with new document: {0}.", Doc.FullName));
                    process(Doc);
                }
                else
                {
                    Console.WriteLine(String.Format("Word initialized with existing document: {0}.", Doc.FullName));
                    process(Doc);
                }
            }
            catch
            {
                Console.WriteLine("No document loaded with word.");
            }
        }

        private void WdEvents2_DocumentOpen(Word.Document Doc)
        {
            process(Doc);
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
