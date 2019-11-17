using System;
using System.IO;
using CMGenerator;
using CMGenerator.Models;
using Serilog;
using Serilog.Core;

namespace cmgenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            var log = Util.GetLog();

            log.Information("Aplicação Iniciada");

            if (!Validar(log))
            {
                Console.WriteLine("Aplicação Finalizado com Erro! Verifique os logs..");
                return;
            }

            bool onlyResults = false;

#if ONLYRESULTS
            onlyResults = true;
#endif

            ProcessWorksheets.Execute(Util.GetWorksheetsDirectory(), Util.GetDirectoryDestination(), log, onlyResults);

            log.Information("Aplicação Finalizada");
        }

        private static bool Validar(Logger log)
        {
            string directory = Util.GetWorksheetsDirectory();
            if (!Directory.Exists(directory))
            {
                log.Warning("Diretório de planilhas não encontrado: {0}", directory);
                return false;
            }

            directory = Util.GetDirectoryDestination();
            if (!Directory.Exists(directory))
                Directory.CreateDirectory(directory);

            return true;
        }
    }
}
