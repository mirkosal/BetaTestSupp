// File: Core/Help.SettingsStore.cs
using System;
using System.IO;
using System.Text;
using System.Collections.Generic;

namespace BetaTestSupp.Core
{
    public sealed class HelpFileSettingsStore : ISettingsStore
    {
        private readonly string _file;
        private readonly Dictionary<string, string?> _kv = new(StringComparer.OrdinalIgnoreCase);

        public HelpFileSettingsStore(string appName, string fileName)
        {
            var dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), appName);
            Directory.CreateDirectory(dir);
            _file = Path.Combine(dir, fileName);
            // load
            if (File.Exists(_file))
                foreach (var line in File.ReadAllLines(_file))
                {
                    var idx = line.IndexOf('=');
                    if (idx <= 0) continue;
                    _kv[line[..idx].Trim()] = line[(idx + 1)..].Trim();
                }
        }

        public string? Get(string key) => _kv.TryGetValue(key, out var v) ? v : null;

        public void Set(string key, string? value)
        {
            _kv[key] = value;
            using var sw = new StreamWriter(_file, false, new UTF8Encoding(false));
            foreach (var kv in _kv)
                sw.WriteLine($"{kv.Key}={kv.Value ?? ""}");
        }
    }
}
