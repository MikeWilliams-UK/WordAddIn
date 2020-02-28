using Microsoft.Office.Tools.Word;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Word.Application app = Application;
            app.DocumentChange += OnDocumentChange;
            app.DocumentBeforeClose += OnDocumentBeforeClose;
        }

        private void OnDocumentBeforeClose(Word.Document doc, ref bool cancel)
        {
            if (Application.Documents.Count > 0)
            {
                doc.ContentControlOnEnter -= OnContentControlEnter;
            }
        }

        private void OnDocumentChange()
        {
            if (Application.Documents.Count > 0)
            {
                Document doc = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveDocument);
                doc.ContentControlOnEnter += OnContentControlEnter;
            }
        }

        private void OnContentControlEnter(Word.ContentControl contentcontrol)
        {
            MessageBox.Show(contentcontrol.ID, contentcontrol.Type.ToString());
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon();
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

        #endregion VSTO generated code
    }
}