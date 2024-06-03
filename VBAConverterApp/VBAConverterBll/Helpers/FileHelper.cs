using System;
using System.IO;

namespace VBAConverterApp.VBAConverterBll.Helpers
{
    public static class FileHelper
    {
        public static string GetRandomFileName()
        {
            string timestamp = DateTime.Now.ToString("yyyyMMddHHmmssfff");
            string randomFilename = $"exportFile_{timestamp}";

            return $"{randomFilename}.xlsx";
        }

        public static string GetMaterialNoLogFileName()
        {
            string timestamp = DateTime.Now.ToString("yyyyMMddHHmmssfff");
            string randomFilename = $"MissingMaterial_{timestamp}";

            return $"{randomFilename}.txt";
        }

        public static string GetMachineConditionLogFileName()
        {
            string timestamp = DateTime.Now.ToString("yyyyMMddHHmmssfff");
            string randomFilename = $"MissingMachineCondition_{timestamp}";

            return $"{randomFilename}.txt";
        }

        public static string GetPathWithRandomFileName()
        {
            string executablePath = AppDomain.CurrentDomain.BaseDirectory;
            string filePath = Path.Combine(executablePath,"Output", GetRandomFileName());

            return filePath;
        }

        public static string GetMissingMaterialNoLogFilePath()
        {
            string executablePath = AppDomain.CurrentDomain.BaseDirectory;
            string filePath = Path.Combine(executablePath, "Output", GetMaterialNoLogFileName());

            return filePath;
        }

        public static string GetMissingMachineConditionLogFilePath()
        {
            string executablePath = AppDomain.CurrentDomain.BaseDirectory;
            string filePath = Path.Combine(executablePath, "Output", GetMachineConditionLogFileName());

            return filePath;
        }

        /// <summary>
        /// Copy file from output template and return new xlsx file as result
        /// </summary>
        /// <returns></returns>
        public static string GetOutputXlsxFile()
        {
            string executablePath = AppDomain.CurrentDomain.BaseDirectory;
            string filePath = Path.Combine(executablePath, "Output", "outputFileTemplate.xlsx");

            string newXlsxFilePath = GetPathWithRandomFileName();
            File.Copy(filePath, newXlsxFilePath);

            return newXlsxFilePath;
        }
    }
}
