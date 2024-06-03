using System.Collections.Generic;

namespace VBAConverterApp.VBAConverterBll.Models
{
    public class BomModel
    {
        public List<MaterialModel> VPowderModels { get; set; } = new List<MaterialModel>();
        public List<MaterialModel> InterPonModels { get; set; } = new List<MaterialModel>();
    }
}
