using System.Collections.Generic;
using VBAConverterApp.VBAConverterBll.Models;

namespace VBAConverterApp.VBAConverterBll.Interfaces
{
    public interface IVBAConverterManager
    {
        List<MaterialModel> ExtractRawMaterial(string filePath, string tabName, int startRow);
    }
}