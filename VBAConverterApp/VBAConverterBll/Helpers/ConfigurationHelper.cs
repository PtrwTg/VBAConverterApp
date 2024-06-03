using Microsoft.Extensions.Configuration;
using System;
using VBAConverterApp.VBAConverterBll.Models;
using ConfigurationBuilder = Microsoft.Extensions.Configuration.ConfigurationBuilder;

namespace VBAConverterApp.VBAConverterBll.Helpers
{
    public static class ConfigurationHelper
    {
        public static VBAConfigModel ReadConfiguration()
        {
            // Build configuration
            var configuration = new ConfigurationBuilder()
                .AddJsonFile("appsettings.json") // Specify the JSON file name
                .Build();

            // Read configuration values
            var result = new VBAConfigModel
            {
                BomFileConfig = new BomFileConfigModel
                {
                    RawTabName = configuration["BomFileConfig:TabName"],
                    RawTabStartLine = Convert.ToInt32(configuration["BomFileConfig:StartLine"]),
                    VPowderScopeRanges = ParseScopeRanges(configuration["BomFileConfig:VPowderScopeRanges"]),
                    InterponScopeRanges = ParseScopeRanges(configuration["BomFileConfig:InterponScopeRanges"]),

                },
                MasterTemplateConfig = new TemplateConfigModel
                {
                    TabName = configuration["MasterTemplateConfig:TabName"],
                    StartLine = Convert.ToInt32(configuration["MasterTemplateConfig:StartLine"])
                },
                MachineConditionConfig = new TemplateConfigModel
                {
                    TabName = configuration["MachineConditionConfig:TabName"],
                    StartLine = Convert.ToInt32(configuration["MachineConditionConfig:StartLine"]),
                    EndLine = Convert.ToInt32(configuration["MachineConditionConfig:EndLine"])
                },
                VPowderTemplateConfig = new TemplateConfigModel
                {
                    TabName = configuration["VPowderTemplateConfig:TabName"],
                    StartLine = Convert.ToInt32(configuration["VPowderTemplateConfig:StartLine"]),
                    EndLine = Convert.ToInt32(configuration["VPowderTemplateConfig:EndLine"])
                }
                ,
                InterponTemplateConfig = new TemplateConfigModel
                {
                    TabName = configuration["InterponTemplateConfig:TabName"],
                    StartLine = Convert.ToInt32(configuration["InterponTemplateConfig:StartLine"]),
                    EndLine = Convert.ToInt32(configuration["InterponTemplateConfig:EndLine"])
                }
            };
            return result;
        }

        private static ExtractScopeRange[] ParseScopeRanges(string ranges)
        {
            string[] rangeStrings = ranges.Split(',');
            ExtractScopeRange[] result = new ExtractScopeRange[rangeStrings.Length];

            for (int i = 0; i < rangeStrings.Length; i++)
            {
                string[] parts = rangeStrings[i].Split('-');
                if (parts.Length == 2)
                {
                    result[i] = new ExtractScopeRange
                    {
                        StartRange = int.Parse(parts[0]),
                        EndRange = int.Parse(parts[1])
                    };
                }
                else
                {
                    throw new FormatException($"Invalid range format: {rangeStrings[i]}");
                }
            }

            return result;
        }
    }
}