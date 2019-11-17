using cmgenerator.Models;
using CMGenerator.Models;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

namespace CMGenerator.Helper
{
    public class ChangeControlReport
    {
        Color colorOnTime = Color.Green;
        Color colorOutOftime = Color.Red;

        const string WORKSHEET_RESUME = "ControleDeMudanças";

        public void Create(ExcelPackage excel, List<Register> registers, List<StockControlReport> stocks)
        {
            Configuration configuration = ConfigurationFactory.Get(null);

            Dictionary<string, string> areaWorksheetNames = new Dictionary<string, string>();

            SetStyles(excel.Workbook);

            var wsResume = excel.Workbook.Worksheets.Add(WORKSHEET_RESUME);

            foreach (var stock in stocks)
            {
                string worksheetName = GetWorksheetName(stock.Area, excel.Workbook.Worksheets);
                areaWorksheetNames.Add(stock.Area, worksheetName);

                var ws = excel.Workbook.Worksheets.Add(worksheetName);
                ws.TabColor = Color.YellowGreen;

                ws.Cells["A1:D1500"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells["A1:D1500"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                ws.Column(1).Width = 25;
                ws.Column(2).Width = 50;
                ws.Column(3).Width = 20;
                ws.Column(4).Width = 20;

                int position = 2;
                position = WriteChangeControlHeader(ws, position, stock);

                position = WriteChangeControlDetails(ws, ++position, registers.Where(x => x.Area.Name.Trim() == stock.Area.Trim() && x.ConclusionDate == DateTime.MinValue).OrderBy(x => x.Number).ToList(), configuration);

                position = WriteChangeControlFooter(ws, position);

                CreateHyperlinkResume(ws, "D2", "Voltar");
            }

            WriteResume(wsResume, areaWorksheetNames, stocks, registers);
        }

        private void WriteResume(ExcelWorksheet ws, Dictionary<string, string> areaWorksheetNames, List<StockControlReport> stocks, List<Register> registers)
        {
            int position = 1;

            var title = ws.Cells["A" + position + ":G" + position];
            title.Style.Font.Bold = true;
            title.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            title.Style.Fill.PatternType = ExcelFillStyle.Solid;
            title.Style.Fill.BackgroundColor.SetColor(Color.LightGray);

            ws.Cells["A" + position].Value = "Area";
            ws.Cells["B" + position].Value = "A Vencer";
            ws.Cells["C" + position].Value = "Atrasados";
            ws.Cells["D" + position].Value = "Total";
            ws.Cells["E" + position].Value = "Encerradas";
            ws.Cells["F" + position].Value = "Encerradas Mês " + (DateTime.Now.Month - 1);
            ws.Cells["G" + position].Value = "Encerradas Mês " + DateTime.Now.Month;

            int totalMonth1 = 0;
            int totalMonth2 = 0;

            foreach (var stock in stocks)
            {
                ws.Cells["A" + ++position].Hyperlink = new ExcelHyperLink(string.Format("'{0}'!A1", areaWorksheetNames[stock.Area]), stock.Area);
                ws.Cells["B" + position].Value = stock.ActionOnTime;
                ws.Cells["C" + position].Value = stock.ActionOutOfTime;
                ws.Cells["D" + position].Value = stock.ActionOnTime + stock.ActionOutOfTime;
                ws.Cells["E" + position].Value = stock.ActionClosed;

                int month1 = GetTotalClosedMonth1(stock.Area, registers);
                ws.Cells["F" + position].Value = month1;
                totalMonth1 += month1;

                int month2 = GetTotalClosedMonth2(stock.Area, registers);
                ws.Cells["G" + position].Value = month1;
                totalMonth2 += month2;
            }

            ws.Cells["A" + ++position].Value = "Total";
            ws.Cells["B" + position].Value = stocks.Sum(x => x.ActionOnTime);
            ws.Cells["C" + position].Value = stocks.Sum(x => x.ActionOutOfTime);
            ws.Cells["D" + position].Value = stocks.Sum(x => x.ActionOnTime + x.ActionOutOfTime);
            ws.Cells["E" + position].Value = stocks.Sum(x => x.ActionClosed);
            ws.Cells["F" + position].Value = totalMonth1;
            ws.Cells["G" + position].Value = totalMonth2;

            ws.Cells["A" + position + ":G" + position].Style.Font.Bold = true;

            SetBorder(ws.Cells["A1:G" + position]);

            ws.Column(1).AutoFit();
            ws.Column(5).AutoFit();
            ws.Column(6).AutoFit();
            ws.Column(7).AutoFit();
        }

        private void SetStyles(ExcelWorkbook workbook)
        {
            var hyperlinkStyle = workbook.Styles.CreateNamedStyle("HyperLink");
            hyperlinkStyle.Style.Font.UnderLine = true;
            hyperlinkStyle.Style.Font.Color.SetColor(Color.Blue);
            hyperlinkStyle.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        }

        private void CreateHyperlinkResume(ExcelWorksheet ws, string range, string name)
        {
            ws.Cells[range].Hyperlink = new ExcelHyperLink(string.Format("'{0}'!A1", WORKSHEET_RESUME), name);
            ws.Cells[range].StyleName = "HyperLink";
        }

        private string GetWorksheetName(string area, ExcelWorksheets worksheets)
        {
            string name = GetWorksheetName(area);
            if (name.Length > 30)
                name = name.Substring(0, 30);

            string aux = name;
            int cont = 1;

            while (worksheets[aux] != null)
                aux = name + cont++;

            return aux;
        }

        private string GetWorksheetName(string area)
        {
            return string.IsNullOrEmpty(area)
                ? "<vazio>"
                : area.Trim().Replace(" ", "").Replace("-", "").Replace("/", "");
        }

        private int WriteChangeControlHeader(ExcelWorksheet ws, int position, StockControlReport stock)
        {
            int startPosition = position;

            var title = ws.Cells["A" + position + ":B" + position];
            title.Merge = true;
            title.Value = "CHANGE CONTROL - " + stock.Area;
            title.Style.Font.Bold = true;
            title.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            title.Style.Fill.PatternType = ExcelFillStyle.Solid;
            title.Style.Fill.BackgroundColor.SetColor(Color.LightGray);


            ws.Cells["A" + ++position].Value = "Status";
            ws.Cells["B" + position].Value = "Quantidade";
            ws.Cells["A" + ++position].Value = "Ação - A vencer";
            ws.Cells["B" + position].Value = stock.ActionOnTime;

            ws.Cells["B" + position].Style.Font.Color.SetColor(colorOnTime);
            ws.Cells["B" + position].Style.Font.Bold = true;
            ws.Cells["A" + ++position].Value = "Ação - Atrasado";
            ws.Cells["B" + position].Value = stock.ActionOutOfTime;

            ws.Cells["B" + position].Style.Font.Color.SetColor(colorOutOftime);
            ws.Cells["B" + position].Style.Font.Bold = true;

            ws.Cells["A" + ++position].Value = "TOTAL";
            ws.Cells["A" + position].Style.Font.Bold = true;
            ws.Cells["B" + position].Value = stock.ActionOnTime + stock.ActionOutOfTime;
            ws.Cells["B" + position].Style.Font.Bold = true;

            SetBorder(ws.Cells["A" + startPosition + ":B" + position]);

            return ++position;
        }

        private int WriteChangeControlDetails(ExcelWorksheet ws, int position, List<Register> registers, Configuration configuration)
        {
            int startPosition = position;
            string columnCM = "A";
            string columnProduct = "B";
            string columnAction= "C";
            string columnPrevisionDate = "D";
            string columnStatus = "E";

            var header = ws.Cells[columnCM + position + ":" + columnStatus + position];
            header.Style.Font.Bold = true;
            header.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            header.Style.Fill.PatternType = ExcelFillStyle.Solid;
            header.Style.Fill.BackgroundColor.SetColor(Color.Gainsboro);
            SetBorder(header);

            ws.Cells[columnCM + position].Value = "CM";
            ws.Cells[columnProduct + position].Value = "Código do Material / Produto";
            ws.Cells[columnAction + position].Value = "Título da ação";
            ws.Cells[columnPrevisionDate + position].Value = "Prazo para Execução";
            ws.Cells[columnStatus + position].Value = "Status";
            ws.Cells[columnStatus + position].AutoFilter = true;

            foreach (var register in registers)
            {
                ws.Cells[columnCM + ++position].Value = register.Number;
                ws.Cells[columnProduct + position].Value = register.Product;
                ws.Cells[columnProduct + position].Style.WrapText = true;
                ws.Cells[columnAction + position].Value = register.Action;
                ws.Cells[columnAction + position].Style.WrapText = true;
                ws.Cells[columnPrevisionDate + position].Value = 
                    register.PrevisionDate.ToString(configuration.DateFormat);
                ws.Cells[columnPrevisionDate + position].Style.Font.Color.SetColor(
                    register.PrevisionDate < DateTime.Now ? colorOutOftime : colorOnTime);
                ws.Cells[columnStatus + position].Value = register.PrevisionDate < DateTime.Now 
                    ? "Atrasado" : "A vencer";
            }

            SetBorder(ws.Cells[columnCM + startPosition + ":" + columnStatus + position]);

            return ++position;
        }

        private int WriteChangeControlFooter(ExcelWorksheet ws, int position)
        {
            position = WriteTable(ws, position, "FSM com pendência da área");

            position = WriteTable(ws, position + 2, "FSM aguardando Comitê Executivo");

            return position;
        }

        private int WriteTable(ExcelWorksheet ws, int position, string titulo)
        {
            position++;

            var title = ws.Cells["A" + ++position + ":C" + position];
            title.Merge = true;
            title.Value = titulo;
            title.Style.Font.Bold = true;
            title.Style.Fill.PatternType = ExcelFillStyle.Solid;
            title.Style.Fill.BackgroundColor.SetColor(Color.YellowGreen);

            title = ws.Cells["A" + ++position + ":C" + position];
            title.Style.Fill.PatternType = ExcelFillStyle.Solid;
            title.Style.Fill.BackgroundColor.SetColor(Color.Gainsboro);
            title.Style.Font.Bold = true;
            ws.Cells["A" + position].Value = "número do FSM";
            ws.Cells["B" + position].Value = "Título";
            ws.Cells["C" + position].Value = "Data de abertura";

            SetBorder(ws.Cells["A" + (position - 1) + ":C" + ++position]);

            return position;
        }

        private void SetBorder(ExcelRange modelTable)
        {
            // Assign borders
            modelTable.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            modelTable.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            modelTable.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            modelTable.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
        }

        #region Months

        bool _calculateMonths = false;
        private DateTime startDateMonth1;
        private DateTime startDateMonth2;
        private DateTime endDateMonth1;
        private DateTime endDateMonth2;

        private int GetTotalClosedMonth1(string area, List<Register> registers)
        {
            if (!_calculateMonths)
                CalculateDatesMonths();

            return registers.Count(x =>
               x.Area.Name == area &&
               x.ConclusionDate >= startDateMonth1 &&
               x.ConclusionDate <= endDateMonth1);
        }

        private int GetTotalClosedMonth2(string area, List<Register> registers)
        {
            if (!_calculateMonths)
                CalculateDatesMonths();

            return registers.Count(x =>
               x.Area.Name == area &&
               x.ConclusionDate >= startDateMonth2 &&
               x.ConclusionDate <= endDateMonth2);
        }

        private void CalculateDatesMonths()
        {
            var month1 = DateTime.Now.AddMonths(-1);
            var month2 = DateTime.Now;

            startDateMonth1 = new DateTime(month1.Year, month1.Month, 1);
            startDateMonth2 = new DateTime(month2.Year, month2.Month, 1);

            endDateMonth1 = new DateTime(month1.Year, month1.Month, DateTime.DaysInMonth(month1.Year, month1.Month));
            endDateMonth2 = new DateTime(month2.Year, month2.Month, DateTime.DaysInMonth(month2.Year, month2.Month));

            _calculateMonths = true;
        }

        #endregion
    }
}
