// File: Core/Help.TemplateRepository.cs
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace BetaTestSupp.Core
{
    public sealed class HelpTemplateRepository : ITemplateRepository
    {
        private readonly string _file;

        public HelpTemplateRepository(string appName, string fileName)
        {
            var dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), appName);
            Directory.CreateDirectory(dir);
            _file = Path.Combine(dir, fileName);
        }

        public IEnumerable<string> List() => ParseAll().Keys;

        public void Save(string name, List<HelpTemplateField> fields)
        {
            var dict = ParseAll();
            dict[name] = fields;
            WriteAll(dict);
        }

        public List<HelpTemplateField>? Load(string name)
        {
            var dict = ParseAll();
            return dict.TryGetValue(name, out var v) ? v : null;
        }

        public void Delete(string name)
        {
            var dict = ParseAll();
            if (dict.Remove(name)) WriteAll(dict);
        }

        private Dictionary<string, List<HelpTemplateField>> ParseAll()
        {
            var dict = new Dictionary<string, List<HelpTemplateField>>(StringComparer.OrdinalIgnoreCase);
            if (!File.Exists(_file)) return dict;

            string? current = null;
            var fields = new List<HelpTemplateField>();

            foreach (var raw in File.ReadAllLines(_file))
            {
                var line = raw.Trim();
                if (string.IsNullOrEmpty(line)) continue;

                if (line.StartsWith("[") && line.EndsWith("]"))
                {
                    if (current != null) dict[current] = fields;
                    current = line[1..^1];
                    fields = new List<HelpTemplateField>();
                    continue;
                }

                if (current == null) continue;
                var idx = line.IndexOf('=');
                if (idx <= 0) continue;
                var key = line[..idx].Trim();
                var val = line[(idx + 1)..].Trim();

                var m = Regex.Match(key, @"^Field(\d+)\.(Name|Pieces)$", RegexOptions.IgnoreCase);
                if (m.Success)
                {
                    int n = int.Parse(m.Groups[1].Value);
                    while (fields.Count <= n) fields.Add(new HelpTemplateField());
                    if (m.Groups[2].Value.Equals("Name", StringComparison.OrdinalIgnoreCase))
                        fields[n].DisplayName = val;
                    else
                        fields[n].Pieces = DeserializePieces(val);
                }
            }
            if (current != null) dict[current] = fields;
            return dict;
        }

        private void WriteAll(Dictionary<string, List<HelpTemplateField>> dict)
        {
            using var sw = new StreamWriter(_file, false, new UTF8Encoding(false));
            foreach (var kv in dict)
            {
                sw.WriteLine($"[{kv.Key}]");
                sw.WriteLine($"Fields={kv.Value.Count}");
                for (int i = 0; i < kv.Value.Count; i++)
                {
                    var f = kv.Value[i];
                    sw.WriteLine($"Field{i}.Name={f.DisplayName}");
                    sw.WriteLine($"Field{i}.Pieces={SerializePieces(f.Pieces)}");
                }
                sw.WriteLine();
            }
        }

        private static string SerializePieces(List<HelpTemplatePiece> pcs)
            => string.Join("|", pcs.Select(p => p.IsToken ? $"T:{p.Value}" : $"L:{p.Value.Replace("|", "\\|")}"));

        private static List<HelpTemplatePiece> DeserializePieces(string s)
        {
            var list = new List<HelpTemplatePiece>();
            foreach (var part in s.Split('|'))
            {
                if (part.StartsWith("T:")) list.Add(new HelpTemplatePiece { IsToken = true, Value = part[2..] });
                else if (part.StartsWith("L:")) list.Add(new HelpTemplatePiece { IsToken = false, Value = part[2..].Replace("\\|", "|") });
            }
            return list;
        }
    }
}
