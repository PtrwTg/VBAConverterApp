using System;
using OfficeOpenXml;
using VBAConverterApp.VBAConverterBll.Constants;

namespace VBAConverterApp.VBAConverterBll.Helpers
{
    public static class ExcelHelper
    {
        public static int FindWorksheetIndex(ExcelPackage excel, string worksheetName)
        {
            try
            {
                int index = -1;

                for (int i = 0; i < excel.Workbook.Worksheets.Count; i++)
                {
                    ExcelWorksheet worksheet = excel.Workbook.Worksheets[i]; // Index is 1-based
                    if (worksheet.Name == worksheetName)
                    {
                        index = i; // Index is 1-based
                        break;
                    }
                }

                return index;
            }
            catch (Exception ex)
            {
                throw new Exception(string.Format(LogMessage.WorksheetTabNotFound, worksheetName) + ex.Message);
            }
            
        }
    }
}