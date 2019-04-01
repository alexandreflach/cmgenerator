using System;
using CsvHelper.Configuration;

namespace cmgenerator.Models
{
    public class Register
    {
        public string Number { get; set; }

        public Area Area { get; set; }

        public string Action { get; set; }

        public DateTime PrevisionDate { get; set; }

        public DateTime ConclusionDate { get; set; }

        public DateTime ExtensionOne { get; set; }

        public DateTime ExtensionTwo { get; set; }

        public DateTime ExtensionThree { get; set; }

        public string Product { get; set; }

        public string Justification { get; set; }

        public string Source { get; set; }

        public override string ToString()
        {
            return Number;
        }
    }

    public sealed class RegisterOnlyResultClassMap : ClassMap<Register>
    {
        public RegisterOnlyResultClassMap()
        {
            Map(m => m.Number).Index(0).Name("CM");
            Map(m => m.Area).Index(1).Name("Área");
            Map(m => m.Action).Index(2).Name("Ação");
            Map(m => m.PrevisionDate).Index(3).Name("Data Prevista");
            Map(m => m.ConclusionDate).Index(4).Name("Data Conclusão");
            Map(m => m.Product).Index(5).Name("Código Descrição");
            Map(m => m.Justification).Index(6).Name("Justificativa");
        }
    }
}
