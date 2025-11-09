// File: Core/Help.UiLogger.cs
using System;
using System.Windows.Forms;

namespace BetaTestSupp.Core
{
    /// <summary>
    /// Logger che scrive su una ListBox dell'interfaccia (thread-safe via Invoke).
    /// </summary>
    public sealed class UiListBoxLogger : IHelpLogger
    {
        private readonly ListBox _list;
        private readonly int _maxItems;

        public UiListBoxLogger(ListBox list, int maxItems = 5000)
        {
            _list = list ?? throw new ArgumentNullException(nameof(list));
            _maxItems = Math.Max(100, maxItems);
        }

        public void Info(string message) => Append($"[INFO]  {message}");
        public void Error(string message) => Append($"[ERRORE] {message}");

        private void Append(string line)
        {
            if (_list.IsDisposed) return;

            void add()
            {
                _list.Items.Add(line);
                // Mantiene la size sotto controllo
                if (_list.Items.Count > _maxItems)
                {
                    int toRemove = _list.Items.Count - _maxItems;
                    for (int i = 0; i < toRemove; i++) _list.Items.RemoveAt(0);
                }
                _list.TopIndex = _list.Items.Count - 1; // autoscroll
            }

            if (_list.InvokeRequired) _list.BeginInvoke((Action)add);
            else add();
        }
    }
}
