// File: Core/Help.FileDialogs.cs
using System;
using System.Windows.Forms;

namespace BetaTestSupp.Core
{
    /// <summary>
    /// Implementazione concreta di IFileDialogService basata su WinForms.
    /// </summary>
    public sealed class WinFileDialogService : IFileDialogService
    {
        private readonly IWin32Window? _owner;

        public WinFileDialogService(IWin32Window? owner = null)
        {
            _owner = owner;
        }

        public string? OpenFile(string filter, string title)
        {
            using var ofd = new OpenFileDialog
            {
                Filter = string.IsNullOrWhiteSpace(filter) ? "Tutti i file|*.*" : filter,
                Title = string.IsNullOrWhiteSpace(title) ? "Seleziona file" : title,
                CheckFileExists = true
            };
            return ofd.ShowDialog(_owner) == DialogResult.OK ? ofd.FileName : null;
        }

        public string? OpenFolder(string title)
        {
            using var fbd = new FolderBrowserDialog
            {
                Description = string.IsNullOrWhiteSpace(title) ? "Seleziona cartella" : title,
                UseDescriptionForTitle = true,
                ShowNewFolderButton = true
            };
            return fbd.ShowDialog(_owner) == DialogResult.OK ? fbd.SelectedPath : null;
        }
    }
}
