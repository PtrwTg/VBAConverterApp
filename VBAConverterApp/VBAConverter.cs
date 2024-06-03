using System;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using VBAConverterApp.VBAConverterBll.Enums;
using VBAConverterApp.VBAConverterBll.Helpers;
using VBAConverterApp.VBAConverterBll.Managers;
using VBAConverterApp.VBAConverterBll.Models;

namespace VBAConverterApp
{
    public partial class VBAConverter : Form
    {
        private readonly VBAConfigModel _config;
        private string _bomFilePath;
        private readonly string _masterFilePath;
        private string _masterRecipeFilePath;
        private OpenFileDialog _openFileDialog1;
        private VBAConverterManager _vbaConverterManager;

        public VBAConverter()
        {
            InitializeComponent();
            InitializeOpenFileDialog();
            _vbaConverterManager = new VBAConverterManager();
            _vbaConverterManager.LogGenerated += Component_LogGenerated;

            _config = ConfigurationHelper.ReadConfiguration();
            string executablePath = AppDomain.CurrentDomain.BaseDirectory;
            _bomFilePath = Path.Combine(executablePath, "SampleFiles\\BOM_UB1P_CDE_Nov.06_ADD Phantom Indicator.xlsx");
            _masterFilePath = Path.Combine(executablePath, "Input\\MasterTemplate.xlsx");
            _masterRecipeFilePath = Path.Combine(executablePath, "SampleFiles\\02 vba SSA_Master recipe rev2_Fresh executed_final_Data for VBA.xlsm");
        }

        private void InitializeOpenFileDialog()
        {
            _openFileDialog1 = new OpenFileDialog
            {
                InitialDirectory = @"C:\",
                Title = "Browse Excel Files",
                CheckFileExists = true,
                CheckPathExists = true,
                DefaultExt = "xlsx",
                Filter = "Excel files (*.xlsx; *.xlsm)|*.xlsx;*.xlsm",
                FilterIndex = 2,
                RestoreDirectory = true,
                ReadOnlyChecked = true,
                ShowReadOnly = true
            };
        }


        private void btnProcess_Click(object sender, EventArgs e)
        {
            Task.Run(Process);
        }

        private void Process()
        {
            // Call the Process method of the manager
            _vbaConverterManager.Process(_bomFilePath, _masterFilePath, _masterRecipeFilePath, _config);
        }

        private void btnBomBrowse_Click(object sender, EventArgs e)
        {
            var button = sender as Button;
            if (button != null && _openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                _bomFilePath = _openFileDialog1.FileName;
                txtBomPath.Text = _bomFilePath;
            }
        }

        private void btnMasterRecipeBrowse_Click(object sender, EventArgs e)
        {
            var button = sender as Button;
            if (button != null && _openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                _masterRecipeFilePath = _openFileDialog1.FileName;
                txtMasterRecipePath.Text = _masterRecipeFilePath;
            }
        }

        private void Component_LogGenerated(object sender, Tuple<string, EnumLogLevel> logInfo)
        {
            // Get log message and severity
            string message = logInfo.Item1;
            EnumLogLevel severity = logInfo.Item2;

            // Set text color based on severity
            Color color = GetLogLevelColor(severity);

            // Update the textbox with the log message and color asynchronously
            BeginInvoke(new Action(() => AppendLog(message, color)));
        }

        // Method to get the color based on log level
        private Color GetLogLevelColor(EnumLogLevel severity)
        {
            switch (severity)
            {
                case EnumLogLevel.Success:
                    return Color.GreenYellow;
                case EnumLogLevel.Info:
                    return Color.White;
                case EnumLogLevel.Warning:
                    return Color.Orange;
                case EnumLogLevel.Error:
                    return Color.Red;
                default:
                    return Color.White;
            }
        }

        // Method to append log message with specified color
        private void AppendLog(string message, Color color)
        {
            if (rtbProcess.InvokeRequired)
            {
                rtbProcess.BeginInvoke(new Action(() =>
                {
                    AppendLog(message, color); // Call the method recursively from the UI thread
                }));
            }
            else
            {
                // Update UI controls directly on the UI thread
                rtbProcess.SelectionStart = rtbProcess.TextLength;
                rtbProcess.SelectionLength = 0;
                rtbProcess.SelectionColor = color;
                rtbProcess.AppendText(message + Environment.NewLine);
                rtbProcess.ScrollToCaret(); // Scroll to the bottom
                rtbProcess.SelectionColor = rtbProcess.ForeColor;
            }
        }
    }
}
