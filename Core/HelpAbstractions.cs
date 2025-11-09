// File: Core/Help.Abstractions.cs
using System.Collections.Generic;

namespace BetaTestSupp.Core
{
    // Logger astratto (UI/Console)
    public interface IHelpLogger
    {
        void Info(string message);
        void Error(string message);
    }

    // Dialog
    public interface IFileDialogService
    {
        string? OpenFile(string filter, string title);
        string? OpenFolder(string title);
    }

    // Settings K/V
    public interface ISettingsStore
    {
        string? Get(string key);
        void Set(string key, string? value);
    }

    // Template (Fase 3)
    public sealed class HelpTemplatePiece
    {
        public bool IsToken { get; set; }        // true = token, false = literal
        public string Value { get; set; } = "";  // token name o testo
        public override string ToString() => IsToken ? "${" + Value + "}" : Value;
    }
    public sealed class HelpTemplateField
    {
        public string DisplayName { get; set; } = "";
        public List<HelpTemplatePiece> Pieces { get; set; } = new();
        public override string ToString() => $"{DisplayName} = {string.Join(" + ", Pieces)}";
    }

    public interface ITemplateRepository
    {
        IEnumerable<string> List();
        void Save(string name, List<HelpTemplateField> fields);
        List<HelpTemplateField>? Load(string name);
        void Delete(string name);
    }

    // Sorgente lista nomi-file (XLSX/CSV/TXT) con supporto Foglio/Colonna
    public interface IListSourceReader
    {
        // Ritorna intestazioni colonna (A: Header, B: Header, …) per il file + foglio (se applicabile)
        List<string> GetHeaders(string path, string? sheetName);
        // Ritorna fogli disponibili (solo XLSX, altrimenti lista vuota o (non applicabile))
        List<string> GetSheets(string path);
        // Estrae i valori dalla colonna scelta
        List<string> ReadValues(string path, string? sheetName, int selectedHeaderIndex0);
    }
}
