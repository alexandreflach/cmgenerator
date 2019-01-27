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
using OfficeOpenXml;

namespace CMGenerator
{
    public class ProcessWorksheets
    {
        internal static void Execute(string directoryWorksheets, string directoryResults, Configuration configuration, Logger log)
        {
            var parserHelper = new WorksheetsParserHelper(configuration, log);

            List<Register> registers = new List<Register>();

            foreach (var file in new DirectoryInfo(directoryWorksheets).GetFiles().OrderBy(x => x.Name))
            {
                Console.WriteLine("Carregando planilha '" + file.Name + "'..");
                try
                {
                    registers.AddRange(parserHelper.Parse(file));
                }
                catch (Exception e)
                {
                    log.Error(e, "Erro ao carregar registros da planilha: {0}", file.Name);
                }
            }

            foreach (var file in new DirectoryInfo(directoryWorksheets).GetFiles().OrderBy(x => x.Name))
            {
                Console.WriteLine("Carregando produtos '" + file.Name + "'..");
                try
                {
                    parserHelper.LoadProducts(file, registers);
                }
                catch (Exception e)
                {
                    log.Error(e, "Erro ao carregar registros produto na planilha: {0}", file.Name);
                }
            }

            Console.WriteLine("Escrevendo resultados.csv");
            GenerateCsvRegisters(directoryResults, registers);

            List<StockControlReport> stocksControl = GetStockControlReport(registers);

            Console.WriteLine("Escrevendo controleacoes.csv");
            GenerateStockControlReport(directoryResults, stocksControl);

            Console.WriteLine("changecontrolreport.xlsx");
            GenerateChangeControlWorkshet(directoryResults, stocksControl, registers, configuration);
        }

        private static void GenerateChangeControlWorkshet(string directoryResults, List<StockControlReport> stocksControl, List<Register> registers, Configuration configuration)
        {
            string fileName = Path.Combine(directoryResults, "changecontrolreport.xlsx");
            if (File.Exists(fileName))
                File.Delete(fileName);

            var fi = new FileInfo(fileName);

            using (var p = new ExcelPackage(fi))
            {
                new ChangeControlReport().Create(p, registers, stocksControl, configuration);
                p.Save();
            }
        }

        private static void GenerateStockControlReport(string directoryResults, List<StockControlReport> registers)
        {
            string fileName = Path.Combine(directoryResults, "controleacoes.csv");
            ExportCsv(registers, fileName);
        }

        private static void GenerateCsvRegisters(string directoryResults, List<Register> registers)
        {
            string fileName = Path.Combine(directoryResults, "resultados.csv");
            ExportCsv(registers, fileName);
        }

        private static List<StockControlReport> GetStockControlReport(List<Register> registers)
        {
            return registers.GroupBy(x => x.Area.Name.Trim())
                            .Select(r =>
                               new StockControlReport
                               {
                                   Area = r.First().Area.Name,
                                   ActionOutOfTime = r.Where(x => x.PrevisionDate.Date <= DateTime.Now.Date && x.ConclusionDate == DateTime.MinValue).Count(),
                                   ActionOnTime = r.Where(x => x.PrevisionDate.Date > DateTime.Now.Date && x.ConclusionDate == DateTime.MinValue).Count(),
                                   ActionClosed = r.Where(x => x.ConclusionDate != DateTime.MinValue && x.ConclusionDate != DateTime.MaxValue).Count(),
                                   ActionCanceled = r.Where(x => x.PrevisionDate == DateTime.MaxValue || x.ConclusionDate == DateTime.MaxValue).Count(),
                                   CountExtensionOne = r.Where(x => x.ExtensionOne != DateTime.MinValue && x.ExtensionOne != DateTime.MaxValue).Count(),
                                   CountExtensionTwo = r.Where(x => x.ExtensionTwo != DateTime.MinValue && x.ExtensionTwo != DateTime.MaxValue).Count(),
                                   CountExtensionThree = r.Where(x => x.ExtensionThree != DateTime.MinValue && x.ExtensionThree != DateTime.MaxValue).Count(),
                                   Total = r.Count()
                               }
                           ).OrderBy(x => x.Area).ToList();
        }

        private static void ExportCsv<T>(List<T> registers, string fileName)
        {
            var stream = new MemoryStream();

            using (var writer = new StreamWriter(stream, Encoding.UTF8))
            {
                var csv = new CsvWriter(writer);
                csv.Configuration.Delimiter = ";";
                csv.WriteHeader<T>();
                csv.NextRecord();
                csv.WriteRecords(registers);
                csv.Flush();

                writer.Flush();
                stream.Flush();

                stream.Position = 0;
                File.WriteAllBytes(fileName, stream.ToArray());
            }
        }
    }
}
