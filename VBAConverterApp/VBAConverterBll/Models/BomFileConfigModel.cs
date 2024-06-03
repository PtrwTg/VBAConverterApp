namespace VBAConverterApp.VBAConverterBll.Models
{
    public class BomFileConfigModel
    {
        public string RawTabName { get; set; }
        public int RawTabStartLine { get; set; }
        public ExtractScopeRange[] VPowderScopeRanges { get; set; }
        public ExtractScopeRange[] InterponScopeRanges { get; set; }
    }

    public class ExtractScopeRange
    {
        public int StartRange { get; set; }
        public int EndRange { get; set; }
    }
}