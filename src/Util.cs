using Serilog;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace CMGenerator
{
    public class Util
    {
        public static Serilog.Core.Logger GetLog()
        {
            return new LoggerConfiguration()
                            .MinimumLevel.Debug()
                            .WriteTo.File("logs\\cmgenerator.txt", rollingInterval: RollingInterval.Day)
                            .CreateLogger();
        }

        public static string GetCurrentDirectory()
        {
            return Directory.GetCurrentDirectory();
        }

        internal static string GetDirectoryDestination()
        {
            return Path.Combine(System.IO.Directory.GetCurrentDirectory(), "resultados");
        }

        public static string GetWorksheetsDirectory()
        {
            return Path.Combine(System.IO.Directory.GetCurrentDirectory(), "planilhas");
        }
    }
}
