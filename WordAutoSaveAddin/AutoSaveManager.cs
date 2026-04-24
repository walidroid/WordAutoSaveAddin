using System;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace WordAutoSaveAddin
{
    /// <summary>
    /// Manages the auto-save timer and document save logic.
    /// Runs on the UI thread via System.Windows.Forms.Timer — no cross-thread marshalling needed.
    /// </summary>
    public class AutoSaveManager : IDisposable
    {
        private const int SaveIntervalMs = 10_000; // 10 seconds — change this to adjust the interval

        private readonly Word.Application _application;
        private readonly Timer _timer;
        private bool _isRunning;
        private bool _disposed;

        public DateTime? LastSavedAt { get; private set; }

        public bool IsRunning => _isRunning;

        public string StatusText
        {
            get
            {
                if (!_isRunning)
                    return "Auto-save: OFF";
                if (LastSavedAt.HasValue)
                    return $"Auto-save: ON  |  Last saved {LastSavedAt.Value:HH:mm:ss}";
                return "Auto-save: ON";
            }
        }

        public AutoSaveManager(Word.Application application)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));

            _timer = new Timer();
            _timer.Interval = SaveIntervalMs;
            _timer.Tick += OnTimerTick;
        }

        public void Start()
        {
            if (_disposed) return;
            _isRunning = true;
            _timer.Start();
        }

        public void Stop()
        {
            _isRunning = false;
            _timer.Stop();
        }

        public void Toggle()
        {
            if (_isRunning)
                Stop();
            else
                Start();
        }

        private void OnTimerTick(object sender, EventArgs e)
        {
            try
            {
                Word.Document doc = _application.ActiveDocument;

                // Skip documents that have never been saved (no path yet)
                if (doc == null || string.IsNullOrEmpty(doc.Path))
                    return;

                // Only save if there are unsaved changes
                if (!doc.Saved)
                {
                    doc.Save();
                    LastSavedAt = DateTime.Now;
                }
            }
            catch (Exception ex)
            {
                // Swallow silently — no popups, no interruption
                System.Diagnostics.Debug.WriteLine($"[AutoSaveManager] Save error: {ex.Message}");
            }
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            _timer.Stop();
            _timer.Tick -= OnTimerTick;
            _timer.Dispose();
        }
    }
}
