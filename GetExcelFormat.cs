using DocumentFormat.OpenXml.Wordprocessing;
using ExcelEf.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;
using DocumentFormat.OpenXml.Office2016.Excel;
using System.Net;
using ExcelEf.Excel;


namespace ExcelEf
{
    internal class GetExcelFormat 
    {
        public static void CreateExcelFile(List<Wells> wells,
            List<WellsDNS> wellsDNS,
            List<WellsHAL> wellsHal,
            List<WellsHAL> wellsTrend,
        DateTime startDate,
        DateTime endDate,
        string filePath
        )
        {
            List<Wells> uniqueWells = UniqieWells(wells);
            //Формируем словарь из ДНСов
            Dictionary<int, string> dnsDictionary = wellsDNS.ToDictionary(d => d.WellId, d => d.DNSName);

            using (var workbook = new XLWorkbook())
            {
                ExcelSheet.GenerateSheetWellop(wells, uniqueWells, dnsDictionary, wellsHal, wellsTrend, startDate, endDate, 1, "Qж (неф)", workbook);
                ExcelSheet.GenerateSheetWellop(wells, uniqueWells, dnsDictionary, wellsHal, wellsTrend, startDate, endDate, 3, "Обв", workbook);
                ExcelSheet.GenerateSheetWellop(wells, uniqueWells, dnsDictionary, wellsHal, wellsTrend, startDate, endDate, 4, "Qн", workbook);
                ExcelSheet.GenerateSheetWellop(wells, uniqueWells, dnsDictionary, wellsHal, wellsTrend, startDate, endDate, 11, "Обв ХАЛ", workbook);
                ExcelSheet.GenerateSheetWellop(wells, uniqueWells, dnsDictionary, wellsHal, wellsTrend, startDate, endDate, 12, "Qж ТМ (неф)", workbook);
                ExcelSheet.GenerateSheetWellop(wells, uniqueWells, dnsDictionary, wellsHal, wellsTrend, startDate, endDate, 213, " Тренд обв ХАЛ", workbook);
                workbook.SaveAs(filePath);
            }
        }
        public static List<Wells> UniqieWells(List<Wells> wells)
        {
            var uniqueWells = new List<Wells>();
            var uniqueId = new HashSet<int>();
            foreach (var well in wells)
            {
                if (uniqueId.Add(well.WellId))
                    uniqueWells.Add(well);
            }
            return uniqueWells;
        }
    }
}
