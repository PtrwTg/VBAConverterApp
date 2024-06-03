using System;
using OfficeOpenXml;

namespace VBAConverterApp.VBAConverterBll.Extensions
{
    public static class ExcelExtensions
    {
        public static string GetStringOrNull(this ExcelRangeBase range)
        {
            string result = range.Value?.ToString().Trim();
            return result == "#NUM!" || result == "#N/A" ? null : result;
        }

        public static decimal? GetDecimalOrNull(this ExcelRangeBase range)
        {
            if (Decimal.TryParse(range.Value?.ToString().Trim(), out decimal result))
            {
                return result;
            }
            else
            {
                return null;
            }
        }
    }
}