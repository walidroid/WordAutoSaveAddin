using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;

namespace WordAutoSaveAddin
{
    [ComVisible(true)]
    public class AutoSaveRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI _ribbon;

        public AutoSaveRibbon()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("WordAutoSaveAddin.AutoSaveRibbon.xml");
        }

        #endregion

        #region Ribbon Callbacks

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            _ribbon = ribbonUI;

            // Start a 2-second UI refresh timer so the status label updates live
            var uiTimer = new System.Windows.Forms.Timer();
            uiTimer.Interval = 2000;
            uiTimer.Tick += (s, e) => _ribbon?.InvalidateControl("statusLabel");
            uiTimer.Start();
        }

        public void OnToggleButton_Click(Office.IRibbonControl control, bool isPressed)
        {
            Globals.ThisAddIn.AutoSaveManager.Toggle();
            _ribbon?.InvalidateControl("statusLabel");
            _ribbon?.InvalidateControl("toggleButton");
        }

        public bool GetToggleButtonPressed(Office.IRibbonControl control)
        {
            return Globals.ThisAddIn.AutoSaveManager.IsRunning;
        }

        public string GetStatusLabel(Office.IRibbonControl control)
        {
            return Globals.ThisAddIn.AutoSaveManager.StatusText;
        }

        #endregion

        #region Helpers

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
