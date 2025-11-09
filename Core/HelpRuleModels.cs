// File: Core/Help.RuleModels.cs
using System;
using System.Collections.Generic;

namespace BetaTestSupp.Core
{
    // Tipi di regola supportati nella prima versione
    public enum RuleKind
    {
        // 1) Valuta una data: esclude se data > oggi (Europe/Rome locale)
        DateAfterToday,

        // 2) Lunghezza massima del testo (es. CF > 16)
        MaxLength,

        // 3) Valore assente nel dataset di Fase 2 (Persone/Corsi) su una singola colonna
        NotPresentInPhase2,

        // 4) Coppia di valori assente nel dataset di Fase 2 (es. CF + PersonNumber, CodCorso + Titolo)
        PairNotPresentInPhase2
    }

    public sealed class RuleDef
    {
        public string Id { get; set; } = Guid.NewGuid().ToString("N");
        public string Name { get; set; } = "Nuova regola";
        public RuleKind Kind { get; set; } = RuleKind.DateAfterToday;

        // Campi riferimento (nomi colonne del file Fase 3)
        public string? Field1 { get; set; }
        public string? Field2 { get; set; }

        // Parametro generico (es. MaxLen)
        public int? IntParam { get; set; }

        // Dataset di riferimento per le regole NotPresentInPhase2/PairNotPresentInPhase2: "Persone" o "Corsi"
        public string? Phase2Dataset { get; set; }
    }

    public interface IRuleRepository
    {
        IEnumerable<RuleDef> List();
        void Save(RuleDef rule);       // upsert (Id)
        void Delete(string id);
    }
}
