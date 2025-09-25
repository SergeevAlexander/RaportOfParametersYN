using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelEf.Excel
{
    internal class ExcelStyle
    {
        internal static void Head(IXLWorksheet worksheet, IXLCell cell, int currentColumn, double Width)
        {
            cell.Style.Font.FontSize = 8;
            cell.Style.Font.FontName = "Arial";
            cell.Style.Font.Bold = true;
            cell.Style.Alignment.WrapText = true;
            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            cell.Style.Border.TopBorder = XLBorderStyleValues.Thin;
            worksheet.Column(currentColumn).Width = Width;
        }
        internal static void Value( IXLCell cell)
        {
            cell.Style.Font.FontSize = 8;
            cell.Style.Font.FontName = "Arial";
            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        }
        internal static void Data(IXLCell cell)
        {
            cell.Style.Font.FontSize = 10;
            cell.Style.Font.FontName = "Arial Cyr";
            cell.Style.Border.TopBorder = XLBorderStyleValues.Thin;
        }
        internal static void CenterRange(IXLRange range)
        {
            range.Style.Font.FontSize = 8;
            range.Style.Font.FontName = "Arial";
            range.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            range.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        }
    }
}
