using System;
using System.IO;
using System.Windows.Forms;

namespace VBAConverterApp.VBAConverterBll.Helpers
{
    public class Logger
    {
        private readonly object _lockObj = new object();
        private readonly string _materialNoLogFilePath;
        private readonly string _machineConditionLogFilePath;
        public string LogPath => _materialNoLogFilePath;

        public Logger()
        {
            _materialNoLogFilePath = FileHelper.GetMissingMaterialNoLogFilePath();
            _machineConditionLogFilePath = FileHelper.GetMissingMachineConditionLogFilePath();
        }

        public void LogMissingMaterialNo(string message)
        {
            lock (_lockObj)
            {
                string logFilePath = _materialNoLogFilePath;
                using (StreamWriter writer = File.AppendText(logFilePath))
                {
                    //string logMessage = $"{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")} - {message}";
                    string logMessage = message;
                    writer.WriteLine(logMessage);
                }
            }
        }

        public void LogMissingMachineCondition(string templateName)
        {
            string logFilePath = _machineConditionLogFilePath;
            using (StreamWriter writer = File.AppendText(logFilePath))
            {
                //string logMessage = $"{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")} - {message}";
                string logMessage = templateName;
                writer.WriteLine(logMessage);
            }
        }
    }
}
