using System;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using VBAConverterApp.VBAConverterBll.Helpers;
using VBAConverterApp.VBAConverterBll.Managers;
using VBAConverterApp.VBAConverterBll.Models;
using Xunit;

namespace VBAConverterAppTest
{
    public class VBAConverterBllTest
    {
        private VBAConfigModel _config;
        private readonly string _bomFilePath;
        private readonly string _masterFilePath;
        private readonly string _masterRecipeFilePath;

        public VBAConverterBllTest()
        {
            _config = ConfigurationHelper.ReadConfiguration();
            string executablePath = AppDomain.CurrentDomain.BaseDirectory;
            _bomFilePath = Path.Combine(executablePath, "SampleFiles\\BOM_UB1P_CDE_Nov.06_ADD Phantom Indicator.xlsx");
            _masterFilePath = Path.Combine(executablePath, "Input\\MasterTemplate.xlsx");
            _masterRecipeFilePath = Path.Combine(executablePath, "SampleFiles\\02 vba SSA_Master recipe rev2_Fresh executed_final_Data for VBA.xlsm");
        }

        [Fact]
        public void RawFileExtractMustSuccess()
        {
            object comObject = null;
            try
            {
                VBAConverterManager converterManager = new VBAConverterManager();
                comObject = converterManager.ExtractRawMaterial(_bomFilePath, _config.BomFileConfig.RawTabName, _config.BomFileConfig.RawTabStartLine);

                Assert.NotNull(comObject);
                Assert.True(comObject.Any());
            }
            finally
            {
                if (comObject != null)
                {
                    Marshal.ReleaseComObject(comObject);
                    comObject = null;
                }
            }
        }

        [Fact]
        public void ProcessMustSuccess()
        {
            VBAConverterManager converterManager = new VBAConverterManager();
            try
            {
                converterManager.Process(_bomFilePath, _masterFilePath, _masterRecipeFilePath, _config);
            }
            catch (Exception ex)
            {
                Assert.True(false, $"Process failed: {ex.Message}");
            }
        }
    }
}
