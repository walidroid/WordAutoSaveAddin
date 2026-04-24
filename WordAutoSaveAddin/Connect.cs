using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Extensibility;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace WordAutoSaveAddin
{
    /// <summary>
    /// Pure COM add-in entry point. Implements IDTExtensibility2 (Word lifecycle)
    /// and IRibbonExtensibility (custom ribbon tab). No VSTO framework required.
    /// </summary>
    [ComVisible(true)]
    [Guid("B2C3D4E5-F678-9012-BCDE-F12345678901")]
    [ProgId("WordAutoSaveAddin.Connect")]
    [ClassInterface(ClassInterfaceType.None)]
    public class Connect : IDTExtensibility2, Office.IRibbonExtensibility
    {
        private Word.Application _application;
        private AutoSaveManager _autoSaveManager;
        private Office.IRibbonUI _ribbon;
        private Timer _uiRefreshTimer;

        // ── IDTExtensibility2 ──────────────────────────────────────────────

        public void OnConnection(object Application, ext_ConnectMode ConnectMode,
                                 object AddInInst, ref Array custom)
        {
            _application = (Word.Application)Application;
            _autoSaveManager = new AutoSaveManager(_application);
            _autoSaveManager.Start();
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            _uiRefreshTimer?.Stop();
            _uiRefreshTimer?.Dispose();
            _autoSaveManager?.Dispose();
        }

        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { }

        // ── IRibbonExtensibility ───────────────────────────────────────────

        public string GetCustomUI(string RibbonID)
        {
            return GetResourceText("WordAutoSaveAddin.AutoSaveRibbon.xml");
        }

        // ── Ribbon callbacks ───────────────────────────────────────────────

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            _ribbon = ribbonUI;

            // Refresh the status label every 2 seconds without interrupting the user
            _uiRefreshTimer = new Timer { Interval = 2000 };
            _uiRefreshTimer.Tick += (s, e) => _ribbon?.InvalidateControl("statusLabel");
            _uiRefreshTimer.Start();
        }

        public void OnToggleButton_Click(Office.IRibbonControl control, bool isPressed)
        {
            _autoSaveManager?.Toggle();
            _ribbon?.InvalidateControl("statusLabel");
            _ribbon?.InvalidateControl("toggleButton");
        }

        public bool GetToggleButtonPressed(Office.IRibbonControl control)
        {
            return _autoSaveManager?.IsRunning ?? false;
        }

        public string GetStatusLabel(Office.IRibbonControl control)
        {
            return _autoSaveManager?.StatusText ?? "Auto-save: OFF";
        }

        // ── Helpers ────────────────────────────────────────────────────────

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            foreach (string name in asm.GetManifestResourceNames())
            {
                if (string.Compare(resourceName, name, StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (var reader = new StreamReader(asm.GetManifestResourceStream(name)))
                        return reader.ReadToEnd();
                }
            }
            return null;
        }
    }
}
