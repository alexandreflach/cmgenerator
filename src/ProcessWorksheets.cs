using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using cmgenerator.Models;
using CMGenerator.Helper;
using CMGenerator.Models;
using CsvHelper;
using Serilog.Core;
using System.Linq;

namespace CMGenerator
{
    public class ProcessWorksheets
    {
        internal static void Execute(string directoryWorksheets, string directoryResults, Configuration configuration, Logger log)
        {
            var parser = new WorksheetsParserHelper(configuration, log);

            List<Register> registers = new List<Register>();

            foreach (var file in new DirectoryInfo(directoryWorksheets).GetFiles())
            {
                Console.WriteLine("Carregando planilha '" + file.Name + "'..");
                try
                {
                    registers.AddRange(parser.Parse(file));
                }
                catch (Exception e)
                {
                    log.Error(e, "Erro ao carregar registros da planilha: {0}", file.Name);
                }
            }

            Console.WriteLine("Escrevendo resultados.csv");
            GenerateCsvRegisters(directoryResults, registers);

            Console.WriteLine("Escrevendo controleacoes.csv");
            GenerateStockControlReport(directoryResults, registers);
        }

        private static void GenerateStockControlReport(string directoryResults, List<Register> registers)
        {
            var list = registers.GroupBy(x => x.Area.Name)
                .Select(r =>
                   new StockControlReport
                   {
                       Area = r.First().Area.Name,
                       ActionOutOfTime = r.Where(x => x.PrevisionDate.Date < DateTime.Now.Date && x.ConclusionDate == DateTime.MinValue).Count(),
                       ActionOnTime = r.Where(x => x.PrevisionDate.Date >= DateTime.Now.Date && x.ConclusionDate == DateTime.MinValue).Count(),
                       ActionClosed = r.Where(x => x.ConclusionDate != DateTime.MinValue && x.ConclusionDate != DateTime.MaxValue).Count(),
                       ActionCanceled = r.Where(x => x.PrevisionDate == DateTime.MaxValue || x.ConclusionDate == DateTime.MaxValue).Count(),
                       Total = r.Count()
                   }
               ).OrderBy(x => x.Area);
                
            var stream = new MemoryStream();

            using (var writer = new StreamWriter(stream, Encoding.UTF8))
            {
                var csv = new CsvWriter(writer);
                csv.Configuration.Delimiter = ";";
                csv.WriteHeader<StockControlReport>();
                csv.NextRecord();
                csv.WriteRecords(list);
                csv.Flush();

                stream.Position = 0;
                File.WriteAllBytes(Path.Combine(directoryResults, "controleacoes.csv"), stream.ToArray());
            }
        }

        private static void GenerateCsvRegisters(string directoryResults, List<Register> registros)
        {
            var stream = new MemoryStream();

            using (var writer = new StreamWriter(stream, Encoding.UTF8))
            {
                var csv = new CsvWriter(writer);
                csv.Configuration.Delimiter = ";";
                csv.WriteHeader<Register>();
                csv.NextRecord();
                csv.WriteRecords(registros);
                csv.Flush();

                stream.Position = 0;
                File.WriteAllBytes(Path.Combine(directoryResults, "resultados.csv"), stream.ToArray());
            }


        }
    }
}
