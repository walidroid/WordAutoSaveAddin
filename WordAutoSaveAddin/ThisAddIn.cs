using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace WordAutoSaveAddin
{
    public partial class ThisAddIn
    {
        private AutoSaveManager _autoSaveManager;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            _autoSaveManager = new AutoSaveManager(this.Application);
            _autoSaveManager.Start();
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            if (_autoSaveManager != null)
            {
                _autoSaveManager.Dispose();
                _autoSaveManager = null;
            }
        }

        public AutoSaveManager AutoSaveManager => _autoSaveManager;

        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new AutoSaveRibbon();
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
