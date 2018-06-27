using cmgenerator.Models;
using CMGenerator.Models;
using OfficeOpenXml;
using Serilog.Core;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

namespace CMGenerator.Helper
{
    public class WorksheetsParserHelper
    {
        public Configuration Configuration { get; }

        public Logger Log { get; }

        public WorksheetsParserHelper(Configuration configuration, Logger log)
        {
            Configuration = configuration;
            Log = log;
        }

        public List<Register> Parse(FileInfo fi)
        {
            List<Register> list = new List<Register>();

            using (var p = new ExcelPackage(fi))
            {
                Log.Information("Carregando planilha: " + fi.Name);

                var ws = GetWorksheet(p);

                LoadColumnPosition(ws);

                for (int i = 2; i < ws.Dimension.End.Row; i++)
                {
                    Register register = GetRegister(ws, i);
                    if (register != null)
                    {
                        register.Source = fi.Name;
                        list.Add(register);
                    }
                }

                Log.Information("Planilha carregada: " + fi.Name);
            }

            return list;
        }

        private Register GetRegister(ExcelWorksheet ws, int rowNumber)
        {
            try
            {
                var number = GetCellValue(ws, rowNumber, Configuration.PositionNumber);
                if (string.IsNullOrEmpty(number)) return null;

                var register = new Register();
                register.Number = number;
                register.Area = AreaHelper.GetArea(GetCellValue(ws, rowNumber, Configuration.PositionResponsibleArea));
                register.Action = Configuration.PositionAction != int.MinValue ? GetCellValue(ws, rowNumber, Configuration.PositionAction) : string.Empty;
                register.PrevisionDate = GetDateCellValue(ws, rowNumber, Configuration.PositionPrevisionDate);
                register.ConclusionDate = GetDateCellValue(ws, rowNumber, Configuration.PositionConclusionDate);
                register.ExtensionOne = GetDateCellValue(ws, rowNumber, Configuration.PositionExtensionOne);
                register.ExtensionTwo = GetDateCellValue(ws, rowNumber, Configuration.PositionExtensionTwo);
                register.ExtensionThree = GetDateCellValue(ws, rowNumber, Configuration.PositionExtensionThree);
                return register;
            }
            catch (Exception e)
            {
                Log.Warning(e, "Erro carregar linha " + rowNumber);
                return null;
            }
        }

        private void LoadColumnPosition(ExcelWorksheet ws)
        {
            Configuration.CleanPosition();

            for (int i = 1; i < ws.Dimension.End.Column; i++)
            {
                var columnName = GetCellValue(ws, 1, i);
                if (Configuration.ColumnNumber.Equals(columnName))
                    Configuration.PositionNumber = i;
                if (Configuration.ColumnResposibleArea.Equals(columnName))
                    Configuration.PositionResponsibleArea = i;
                if (Configuration.ColumnAction.Equals(columnName))
                    Configuration.PositionAction = i;
                if (Configuration.ColumnPrevisionDate.Equals(columnName))
                    Configuration.PositionPrevisionDate = i;
                if (Configuration.ColumnConclusionDate.Equals(columnName))
                    Configuration.PositionConclusionDate = i;
                if (Configuration.ColumnExtensionOne.Equals(columnName))
                    Configuration.PositionExtensionOne = i;
                if (Configuration.ColumnExtensionTwo.Equals(columnName))
                    Configuration.PositionExtensionTwo = i;
                if (Configuration.ColumnExtensionThree.Equals(columnName))
                    Configuration.PositionExtensionThree = i;
            }

            Configuration.ValidatedPosition();
        }

        private string GetCellValue(ExcelWorksheet ws, int rowNumber, int columnNumber)
        {
            var value = ws.Cells[rowNumber, columnNumber].Value;
            return value != null ? value.ToString() : string.Empty;
        }

        private DateTime GetDateCellValue(ExcelWorksheet ws, int rowNumber, int columnNumber)
        {
            var value = ws.Cells[rowNumber, columnNumber].Value;

            if (value == null) return DateTime.MinValue;

            if (value != null && value is DateTime)
                return (DateTime)value;

            string valueText = value.ToString();

            if (valueText.Trim().ToUpper().Equals("NA")) return DateTime.MaxValue;

            try
            {
                return ParseDate(value.ToString());
            }
            catch
            {
                return DateTime.MaxValue;
            }
        }

        private DateTime ParseDate(string date)
        {
            return DateTime.ParseExact(date, Configuration.DateFormat, CultureInfo.InvariantCulture);
        }

        private ExcelWorksheet GetWorksheet(ExcelPackage p)
        {
            foreach (var w in p.Workbook.Worksheets)
                if (w.Name.Contains(Configuration.WorksheetName)) return w;

            throw new Exception("Não encontrada planilha, verifique se existe a planilha com nome: " + Configuration.WorksheetName);
        }
    }
}
