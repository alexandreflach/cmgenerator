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

        internal void LoadProductsAndJustification(FileInfo file, List<Register> registers, bool fsm)
        {
            using (var p = new ExcelPackage(file))
            {
                try
                {
                    var ws = GetWorksheetProduct(p, fsm);
                    LoadColumnPosition(ws, false);
                    if(Configuration.PositionNumber == int.MinValue || Configuration.PositionProduct == int.MinValue)
                    {
                        Log.Warning(string.Format("Colunas ({0} e {1}) não encontradas Planilha ({2}) arquivo: {3}", 
                            Configuration.ColumnNumber, Configuration.ColumnProduct, ws.Name, file.Name));
                        return;
                    }

                    for (int i = 2; i < ws.Dimension.End.Row; i++)
                    {
                        var number = GetCellValue(ws, i, Configuration.PositionNumber);
                        if (string.IsNullOrEmpty(number) || !registers.Exists(x => x.Number == number))
                            continue;

                        var product = GetCellValue(ws, i, Configuration.PositionProduct);
                        var justification = GetCellValue(ws, i, Configuration.PositionJustification);

                        if (string.IsNullOrEmpty(product) && string.IsNullOrEmpty(justification)) continue;

                        foreach (var r in registers.FindAll(x => x.Number == number))
                        {
                            if (!string.IsNullOrEmpty(product))
                                r.Product = product;
                            if (!string.IsNullOrEmpty(justification))
                                r.Justification = justification;
                        }
                    }
                }
                catch (FileNotFoundException e)
                {
                    Log.Warning(string.Format("Planilha ({0}) não encontrada em {1}", e.Message, file.Name));
                }
            }
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

        private void LoadColumnPosition(ExcelWorksheet ws, bool validar = true)
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
                if (Configuration.ColumnProduct.Equals(columnName))
                    Configuration.PositionProduct = i;
                if (Configuration.ColumnJustification.Equals(columnName))
                    Configuration.PositionJustification = i;
            }

            if(validar)
                Configuration.ValidatedPosition();
        }

        private string GetCellValue(ExcelWorksheet ws, int rowNumber, int columnNumber)
        {
            var value = ws.Cells[rowNumber, columnNumber].Value;
            return value != null ? value.ToString().Trim() : string.Empty;
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

        private ExcelWorksheet GetWorksheetProduct(ExcelPackage p, bool fsm)
        {
            foreach (var w in p.Workbook.Worksheets)
                if (w.Name.Contains(fsm ? Configuration.WorksheetFsmProductName : Configuration.WorksheetProductName)) return w;

            throw new FileNotFoundException(fsm ? Configuration.WorksheetFsmProductName : Configuration.WorksheetProductName);
        }
    }
}
