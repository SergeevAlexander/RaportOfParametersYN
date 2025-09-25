using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using ExcelEf.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading.Tasks;

namespace ExcelEf.Excel
{
    internal class ExcelSheet
    {
        public static void GenerateSheetWellop(
                List<Wells> records,
                List<Wells> uniqueWells,
                Dictionary<int, string> dnsDictionary,
                List<WellsHAL> wellsHal,
                    List<WellsHAL> wellsTrend,
                    DateTime startDate,
                    DateTime endDate,
                    int wellMeasureTypeId,
                    string nameSheets,
                    XLWorkbook workbook
                    )
        {
            var worksheet = workbook.Worksheets.Add(nameSheets);

            int wellopCount = uniqueWells.Count;

            //Формируем шапку в экселе
            SheetWellopHead(worksheet, startDate, endDate, wellopCount);

            //Загружаем данные в таблицу
            SheetWellopData(worksheet, uniqueWells, wellMeasureTypeId, dnsDictionary, startDate, endDate, records, wellsHal, wellsTrend);

        }

        private static void SheetWellopHead(
            IXLWorksheet worksheet,
            DateTime startDate,
            DateTime endDate,
            int wellopCount)
        {
            int currentRow = 2;
            int currentColumn = 2;

            worksheet.Row(2).Height = 45;
            worksheet.Column(1).Width = 1.57;
            //Название колонок
            var headNames = new Dictionary<double, string>()
            {
                {5,  "№ п/п"}, {13,  "Местор-ие"}, {4.71, "Код мест-я"}, {5.57, "№ скважины"}, {5.71, "№ куста"}, {8, "ЦДНГ"}, {7.14, "№ бригады"},
                {7.71, "Значение по ТР"}, {12.71, "Текущий способ эксплуатации"}, {21.57, "ДНС"}
            };
            foreach (var headName in headNames)
            {
                HeadValue(worksheet, currentColumn++, currentRow, headName.Key, headName.Value);
            }
            //Шапка даты
            for (var date = startDate; date <= endDate; date = date.AddDays(1))
            {
                HeadValue(worksheet, currentColumn, currentRow, 4.29, date.ToString("dd.MM.yyyy"));
                currentColumn++;
            }
            currentColumn++;
            HeadValue(worksheet, currentColumn, currentRow, 4.29, "Последний");
            //Формулы для количества замеров
            ItogHead(worksheet);
            currentColumn = 11;
            for (var date = startDate; date <= endDate; date = date.AddDays(1))
            {
                var range = worksheet.Range(
                worksheet.Cell(11, currentColumn), worksheet.Cell(wellopCount + 11, currentColumn));

                worksheet.Cell(6, currentColumn).FormulaA1 = $"=SUBTOTAL(2, {range.RangeAddress})";
                ExcelStyle.Value(worksheet.Cell(currentRow, currentColumn));
                worksheet.Cell(7, currentColumn).FormulaA1 = $"=SUBTOTAL(2, {range.RangeAddress})";
                ExcelStyle.Value(worksheet.Cell(currentRow, currentColumn));
            }
        }

        private static void HeadValue(
            IXLWorksheet worksheet,
            int currentColumn,
            int currentRow,
            double width,
            string value)
        {
            ExcelStyle.Head(worksheet, worksheet.Cell(currentRow, currentColumn), currentColumn, width);
            worksheet.Cell(currentRow, currentColumn).Value = value;
        }
        private static void ItogHead(IXLWorksheet worksheet)
        {
            worksheet.Rows(3, 5).Hide();
            var itogRange = worksheet.Range("F6:K6");
            itogRange.Merge().Value = "Итоги по всему фонду";
            ExcelStyle.CenterRange(itogRange);
            itogRange = worksheet.Range("F7:K7");
            itogRange.Merge().Value = "Cуточные итоги";
            ExcelStyle.CenterRange(itogRange);
            itogRange = worksheet.Range("F8:K8");
            itogRange.Merge().Value = "Недельные итоги";
            ExcelStyle.CenterRange(itogRange);
            itogRange = worksheet.Range("F9:K9");
            itogRange.Merge().Value = "Месячные итоги";
            ExcelStyle.CenterRange(itogRange);
            worksheet.Range("B10:K10").SetAutoFilter();
            worksheet.Range("B6:K10").Style.Border.TopBorder = XLBorderStyleValues.Thin;
        }
        private static void SheetWellopData(
            IXLWorksheet worksheet,
            List<Wells> uniqueWells,
            int wellMeasureTypeId,
            Dictionary<int, string> dnsLookup,
            DateTime startDate,
            DateTime endDate,
            List<Wells> records,
            List<WellsHAL> wellsHal,
            List<WellsHAL> wellsTrend
            )
        {
            //Формируем список значений, которые должны выводиться в зависимости от параметрам wellMeasureTypeId
            List<Wells> selectedWells = SelectedWells(records, wellMeasureTypeId, wellsHal, wellsTrend);
            //Формируем словарь с значениями
            var measureWells = selectedWells
                .GroupBy(r => (r.WellId, r.MeasureDate))
                .ToDictionary(
                    g => g.Key,
                    g => g.OrderByDescending(r => r.MeasureDate).First().Value
                    );
            //Формируем Словарь с последними значениями
            Dictionary<int, double> latestValues = selectedWells
                .GroupBy(d => d.WellId)
                .ToDictionary(g => g.Key,
                g => g.OrderByDescending(d => d.MeasureDate).First().Value);

            var currentRow = 11;
            foreach (var well in uniqueWells.OrderBy(w => w.WellId))
            {
                var currentColumn = 2;
                DataNumber(worksheet, currentColumn++, currentRow);
                DataValue(worksheet, currentColumn++, currentRow, well.FieldName);
                DataValue(worksheet, currentColumn++, currentRow, well.FieldCode);
                DataValue(worksheet, currentColumn++, currentRow, well.WellName);
                DataValue(worksheet, currentColumn++, currentRow, well.WellCluster);
                DataValue(worksheet, currentColumn++, currentRow, well.ShopName);
                DataValue(worksheet, currentColumn++, currentRow, well.TeamNo);
                //Выбор параметра ТР
                ExcelStyle.Data(worksheet.Cell(currentRow, currentColumn));
                DataTRValue(worksheet, wellMeasureTypeId, currentColumn++, currentRow, well.TRQg, well.TRObv, well.TRQn);
                DataValue(worksheet, currentColumn++, currentRow, well.TekSposob);

                if (dnsLookup.TryGetValue(well.WellId, out string dnsName))
                {
                    DataValue(worksheet, currentColumn++, currentRow, $"{dnsName} ({well.ShopName})");
                }
                else
                {
                    DataValue(worksheet, currentColumn++, currentRow, "");
                }
                for (var date = startDate; date <= endDate; date = date.AddDays(1))
                {
                    var key = (well.WellId, date.Date);
                    if (measureWells.TryGetValue(key, out double amount))
                    {
                        DataValue(worksheet, currentColumn, currentRow, Math.Round(amount, 2));
                    }
                    else
                    {
                        ExcelStyle.Value(worksheet.Cell(currentRow, currentColumn));
                        worksheet.Cell(currentRow, currentColumn).Value = "";
                    }
                    currentColumn++;
                }
                currentColumn++;
                if (latestValues.TryGetValue(well.WellId, out double Value))
                {
                    DataValue(worksheet, currentColumn, currentRow, Value);
                }
                currentRow++;
            }
        }
        private static void DataValue(
            IXLWorksheet worksheet,
            int currentColumn,
            int currentRow,
            string value)
        {
            ExcelStyle.Data(worksheet.Cell(currentRow, currentColumn));
            worksheet.Cell(currentRow, currentColumn).Value = value;
        }
        private static void DataValue(
            IXLWorksheet worksheet,
            int currentColumn,
            int currentRow,
            int value)
        {
            ExcelStyle.Data(worksheet.Cell(currentRow, currentColumn));
            worksheet.Cell(currentRow, currentColumn).Value = value;
        }
        private static void DataValue(
            IXLWorksheet worksheet,
            int currentColumn,
            int currentRow,
            double value)
        {
            ExcelStyle.Value(worksheet.Cell(currentRow, currentColumn));
            worksheet.Cell(currentRow, currentColumn).Value = value;
        }
        private static void DataNumber(
            IXLWorksheet worksheet,
            int currentColumn,
            int currentRow)
        {
            ExcelStyle.Data(worksheet.Cell(currentRow, currentColumn));
            worksheet.Cell(currentRow, currentColumn).Value = currentRow - 10;
        }
        private static void DataTRValue(
            IXLWorksheet worksheet,
            int wellMeasureTypeId,
            int currentColumn,
            int currentRow,
            double TRQg,
            double TRObv,
            double TRQn)
        {
            switch (wellMeasureTypeId)
            {
                case 1:
                    worksheet.Cell(currentRow, currentColumn).Value = TRQg;
                    break;
                case 3:
                    worksheet.Cell(currentRow, currentColumn).Value = TRObv;
                    break;
                case 4:
                    worksheet.Cell(currentRow, currentColumn).Value = Math.Round(TRQn, 2);
                    break;
                default:
                    worksheet.Cell(currentRow, currentColumn).Value = 0;
                    break;
            }
        }
        private static List<Wells> SelectedWells(List<Wells> wells, int wellMeasureTypeId, List<WellsHAL> wellsHAL,
           List<WellsHAL> wellsTrend)
        {
            var selectedWells = new List<Wells>();
            switch (wellMeasureTypeId)
            {
                case 11:
                    var lookupDict = wellsHAL
                    .GroupBy(wh => (wh.WellId, wh.MeasureDate))
                    .ToDictionary(g => g.Key, g => g.First().Value);
                    foreach (var well in wells)
                    {
                        var key = (well.WellId, well.MeasureDate);
                        if (lookupDict.TryGetValue(key, out double value))
                        {
                            well.Value = value;
                            selectedWells.Add(well);
                        }
                    }
                    return selectedWells;
                case 213:
                    var lookupDictTrend = wellsTrend
                .GroupBy(wh => (wh.WellId, wh.MeasureDate))
                .ToDictionary(g => g.Key, g => g.First().Value);
                    foreach (var well in wells)
                    {
                        var key = (well.WellId, well.MeasureDate);
                        if (lookupDictTrend.TryGetValue(key, out double value))
                        {
                            well.Value = value;
                            selectedWells.Add(well);
                        }
                    }
                    return selectedWells;
                case 1:
                    foreach (var well in wells)
                    {
                        if (well.MeasureTypeId == wellMeasureTypeId)
                            selectedWells.Add(well);
                    }
                    return selectedWells;
                case 3: goto case 1;
                case 4: goto case 1;
                case 12: goto case 1;
                default: return selectedWells;
            }
        }

    }
}
