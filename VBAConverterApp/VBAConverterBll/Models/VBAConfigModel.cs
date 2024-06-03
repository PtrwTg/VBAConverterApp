namespace VBAConverterApp.VBAConverterBll.Models
{
    public class VBAConfigModel
    {
        public BomFileConfigModel BomFileConfig { get; set; }
        public TemplateConfigModel MasterTemplateConfig { get; set; }
        public TemplateConfigModel MachineConditionConfig { get; set; }
        public TemplateConfigModel VPowderTemplateConfig { get; set; }
        public TemplateConfigModel InterponTemplateConfig { get; set; }
    }

}