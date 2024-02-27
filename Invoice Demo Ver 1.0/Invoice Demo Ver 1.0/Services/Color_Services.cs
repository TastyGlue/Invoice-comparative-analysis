using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Invoice_Demo_Ver_1._0.Services
{
    public static class Color_Services
    {
        //public static Color highlightColor = Color.Yellow;

        public static void HighlightNRAObject(ExcelWorksheet worksheet, int row, int col, Color highlightColor)
        {
            for (int colToColor = col; colToColor < col+15; colToColor++)
            {
                worksheet.Cells[row, colToColor].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                worksheet.Cells[row, colToColor].Style.Fill.BackgroundColor.SetColor(highlightColor);
            }
        }

        public static void HighlightAzhurObject(ExcelWorksheet worksheet, int row, int col, Color highlightColor)
        {
            for (int colToColor = col; colToColor < col+9; colToColor++)
            {
                worksheet.Cells[row, colToColor].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                worksheet.Cells[row, colToColor].Style.Fill.BackgroundColor.SetColor(highlightColor);
            }
        }

        public static void HighlightCell(ExcelWorksheet worksheet, int row, int col, Color highlightColor)
        {
            worksheet.Cells[row, col].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[row, col].Style.Fill.BackgroundColor.SetColor(highlightColor);
        }
    }
}
