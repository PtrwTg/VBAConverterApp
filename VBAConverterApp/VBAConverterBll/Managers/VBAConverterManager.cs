using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using VBAConverterApp.VBAConverterBll.Constants;
using VBAConverterApp.VBAConverterBll.Enums;
using VBAConverterApp.VBAConverterBll.Extensions;
using VBAConverterApp.VBAConverterBll.Helpers;
using VBAConverterApp.VBAConverterBll.Interfaces;
using VBAConverterApp.VBAConverterBll.Models;

namespace VBAConverterApp.VBAConverterBll.Managers
{
    public class VBAConverterManager : IVBAConverterManager
    {
        private readonly string[] _updateComponents = { ComponentNumbers.Number0050, ComponentNumbers.Number0060, ComponentNumbers.Number0070, ComponentNumbers.Number0080, ComponentNumbers.Number0090, ComponentNumbers.Number0100, ComponentNumbers.Number0110, ComponentNumbers.Number0040 };

        private readonly Logger _logger;

        // Define an event to handle logging
        public event EventHandler<Tuple<string, EnumLogLevel>> LogGenerated;


        public VBAConverterManager()
        {
            _logger = new Logger();
        }

        /// <summary>
        /// Read Raw material to object
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="tabName"></param>
        /// <param name="startRow"></param>
        /// <returns></returns>
        public List<MaterialModel> ExtractRawMaterial(string filePath, string tabName, int startRow)
        {
            List<MaterialModel> materials = new List<MaterialModel>();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage excel = new ExcelPackage(new FileInfo(filePath)))
            {
                // find worksheet index by tab name
                int worksheetIndex = ExcelHelper.FindWorksheetIndex(excel, tabName);
                ExcelWorksheet worksheet = excel.Workbook.Worksheets[worksheetIndex];

                int rowCount = worksheet.Dimension.Rows;
                if (rowCount > 0)
                {
                    for (int row = startRow; row <= rowCount; row++) // Assuming the first row is the header
                    {
                        MaterialModel material = new MaterialModel
                        {
                            Plnt = worksheet.Cells[row, 1].GetStringOrNull(),
                            Material = Convert.ToInt32(worksheet.Cells[row, 2].Value),
                            PhantomIndicator = worksheet.Cells[row, 3].GetStringOrNull(),
                            AltBOM = worksheet.Cells[row, 4].GetStringOrNull(),
                            MaterialNumber = worksheet.Cells[row, 5].GetStringOrNull(),
                            BOMUsage = worksheet.Cells[row, 6].GetStringOrNull(),
                            LabOrOffice = worksheet.Cells[row, 7].GetStringOrNull(),
                            BaseQty = worksheet.Cells[row, 8].GetStringOrNull(),
                            Unit = worksheet.Cells[row, 9].GetStringOrNull(),
                            ICt = worksheet.Cells[row, 10].GetStringOrNull(),
                            Item = worksheet.Cells[row, 11].GetStringOrNull(),
                            Component = worksheet.Cells[row, 12].GetStringOrNull(),
                            Un = worksheet.Cells[row, 13].GetStringOrNull(),
                            Quantity1 = worksheet.Cells[row, 14].GetDecimalOrNull(),
                            Quantity1Unit = worksheet.Cells[row, 15].GetStringOrNull(),
                            BOMComponent = worksheet.Cells[row, 16].GetStringOrNull(),
                            ItemTextLine1 = worksheet.Cells[row, 17].GetStringOrNull(),
                            ChangeNo = worksheet.Cells[row, 18].GetStringOrNull(),
                            SPT = worksheet.Cells[row, 19].GetStringOrNull(),
                            BOM = worksheet.Cells[row, 20].GetStringOrNull(),
                            ChangedOn = worksheet.Cells[row, 21].GetStringOrNull(),
                            SortString = worksheet.Cells[row, 22].GetStringOrNull(),
                            BOMSt = worksheet.Cells[row, 23].GetStringOrNull(),
                            DID = worksheet.Cells[row, 24].GetStringOrNull(),
                            DeID = worksheet.Cells[row, 25].GetStringOrNull(),
                            ProdS = worksheet.Cells[row, 26].GetStringOrNull(),
                            MRPC = worksheet.Cells[row, 27].GetStringOrNull(),
                            MS = worksheet.Cells[row, 28].GetStringOrNull(),
                            Pl = worksheet.Cells[row, 29].GetStringOrNull(),
                            IsLoc = worksheet.Cells[row, 30].GetStringOrNull(),
                            SupplyArea = worksheet.Cells[row, 31].GetStringOrNull()
                        };

                        materials.Add(material);
                    }
                }
            }

            return materials;
        }

        public List<MaterialModel> ExtractMaterialRange(List<MaterialModel> rawMaterialModels, int startRange, int endRange)
        {
            return rawMaterialModels.Where(x => x.Material >= startRange && x.Material <= endRange).ToList();
        }

        /// <summary>
        /// Read Master template file to object
        /// </summary>
        /// <param name="masterFilePath"></param>
        /// <param name="masterTemplateConfig"></param>
        /// <returns></returns>
        public List<MasterDataModel> ReadMasterTemplate(string masterFilePath, TemplateConfigModel masterTemplateConfig)
        {
            List<MasterDataModel> masterDataList = new List<MasterDataModel>();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            // Load the Excel file
            using (ExcelPackage package = new ExcelPackage(new FileInfo(masterFilePath)))
            {
                // find worksheet index by tab name
                int worksheetIndex = ExcelHelper.FindWorksheetIndex(package, masterTemplateConfig.TabName);
                ExcelWorksheet worksheet = package.Workbook.Worksheets[worksheetIndex]; // Assuming data is in the first worksheet

                // Iterate through rows in the Excel file
                for (int row = masterTemplateConfig.StartLine; row <= worksheet.Dimension.Rows; row++) // Start from row 2 to skip header
                {
                    // Create a new MasterDataModel object for each row
                    MasterDataModel masterData = new MasterDataModel();

                    // Map each cell value to the corresponding property in MasterDataModel
                    masterData.FieldName = worksheet.Cells[row, 1].GetStringOrNull();
                    masterData.SK1 = worksheet.Cells[row, 2].GetStringOrNull();
                    masterData.STLNR = worksheet.Cells[row, 3].GetStringOrNull();
                    masterData.MATNR = worksheet.Cells[row, 4].GetStringOrNull();
                    masterData.WERKS = worksheet.Cells[row, 5].GetStringOrNull();
                    masterData.STLAN = worksheet.Cells[row, 6].GetStringOrNull();
                    masterData.STLAL = worksheet.Cells[row, 7].GetStringOrNull();
                    masterData.BMENG = worksheet.Cells[row, 8].GetStringOrNull();
                    masterData.BMEIN = worksheet.Cells[row, 9].GetStringOrNull();
                    masterData.LABOR = worksheet.Cells[row, 10].GetStringOrNull();
                    masterData.STLKN = worksheet.Cells[row, 11].GetStringOrNull();
                    masterData.ALPOS = worksheet.Cells[row, 12].GetStringOrNull();
                    masterData.AUSCH = worksheet.Cells[row, 13].GetStringOrNull();
                    masterData.AVOAU = worksheet.Cells[row, 14].GetStringOrNull();
                    masterData.BEIKZ = worksheet.Cells[row, 15].GetStringOrNull();
                    masterData.CADPO = worksheet.Cells[row, 16].GetStringOrNull();
                    masterData.EKGRP = worksheet.Cells[row, 17].GetStringOrNull();
                    masterData.FMENG = worksheet.Cells[row, 18].GetStringOrNull();
                    masterData.IDNRK = worksheet.Cells[row, 19].GetStringOrNull();
                    masterData.LIFNR = worksheet.Cells[row, 20].GetStringOrNull();
                    masterData.LIFZT = worksheet.Cells[row, 21].GetStringOrNull();
                    masterData.MATKL = worksheet.Cells[row, 22].GetStringOrNull();
                    masterData.MEINS = worksheet.Cells[row, 23].GetStringOrNull();
                    masterData.MENGE = worksheet.Cells[row, 24].GetDecimalOrNull();
                    masterData.NETAU = worksheet.Cells[row, 25].GetStringOrNull();
                    masterData.NFMAT = worksheet.Cells[row, 26].GetStringOrNull();
                    masterData.PEINH = worksheet.Cells[row, 27].GetStringOrNull();
                    masterData.POSNR = worksheet.Cells[row, 28].GetStringOrNull();
                    masterData.POSTP = worksheet.Cells[row, 29].GetStringOrNull();
                    masterData.PREIS = worksheet.Cells[row, 30].GetStringOrNull();
                    masterData.POTX1 = worksheet.Cells[row, 31].GetStringOrNull();
                    masterData.SK2 = worksheet.Cells[row, 32].GetStringOrNull();
                    masterData.SK3 = worksheet.Cells[row, 33].GetStringOrNull();
                    masterData.SK4 = worksheet.Cells[row, 34].GetStringOrNull();
                    masterData.POTX2 = worksheet.Cells[row, 35].GetStringOrNull();
                    masterData.PSWRK = worksheet.Cells[row, 36].GetStringOrNull();
                    masterData.REKRS = worksheet.Cells[row, 37].GetStringOrNull();
                    masterData.RFORM = worksheet.Cells[row, 38].GetStringOrNull();
                    masterData.ROANZ = worksheet.Cells[row, 39].GetStringOrNull();
                    masterData.ROMEI = worksheet.Cells[row, 40].GetStringOrNull();
                    masterData.ROMEN = worksheet.Cells[row, 41].GetStringOrNull();
                    masterData.ROMS1 = worksheet.Cells[row, 42].GetStringOrNull();
                    masterData.ROMS2 = worksheet.Cells[row, 43].GetStringOrNull();
                    masterData.ROMS3 = worksheet.Cells[row, 44].GetStringOrNull();
                    masterData.RVREL = worksheet.Cells[row, 45].GetStringOrNull();
                    masterData.SAKTO = worksheet.Cells[row, 46].GetStringOrNull();
                    masterData.SANFE = worksheet.Cells[row, 47].GetStringOrNull();
                    masterData.SANKA = worksheet.Cells[row, 48].GetStringOrNull();
                    masterData.SANKO = worksheet.Cells[row, 49].GetStringOrNull();
                    masterData.SANVS = worksheet.Cells[row, 50].GetStringOrNull();
                    masterData.SCHGT = worksheet.Cells[row, 51].GetStringOrNull();
                    masterData.SORTF = worksheet.Cells[row, 52].GetStringOrNull();
                    masterData.PLNNR = worksheet.Cells[row, 53].GetStringOrNull();
                    masterData.PLNAL = worksheet.Cells[row, 54].GetStringOrNull();
                    masterData.ProductVersion = worksheet.Cells[row, 55].GetStringOrNull();
                    masterData.SK6 = worksheet.Cells[row, 56].GetStringOrNull();
                    masterData.VERTI = worksheet.Cells[row, 57].GetStringOrNull();
                    masterData.WEBAZ = worksheet.Cells[row, 58].GetStringOrNull();
                    masterData.WAERS = worksheet.Cells[row, 59].GetStringOrNull();
                    masterData.DOKAR = worksheet.Cells[row, 60].GetStringOrNull();
                    masterData.DOKNR = worksheet.Cells[row, 61].GetStringOrNull();
                    masterData.DOKVR = worksheet.Cells[row, 62].GetStringOrNull();
                    masterData.DOKTL = worksheet.Cells[row, 63].GetStringOrNull();
                    masterData.EWAHR = worksheet.Cells[row, 64].GetStringOrNull();
                    masterData.EKORG = worksheet.Cells[row, 65].GetStringOrNull();
                    masterData.LGORT = worksheet.Cells[row, 66].GetStringOrNull();
                    masterData.CLASS = worksheet.Cells[row, 67].GetStringOrNull();
                    masterData.KLART = worksheet.Cells[row, 68].GetStringOrNull();
                    masterData.POTPR = worksheet.Cells[row, 69].GetStringOrNull();
                    masterData.ALPGR = worksheet.Cells[row, 70].GetStringOrNull();
                    masterData.ALPST = worksheet.Cells[row, 71].GetStringOrNull();
                    masterData.ALPRF = worksheet.Cells[row, 72].GetStringOrNull();
                    masterData.DSPST = worksheet.Cells[row, 73].GetStringOrNull();
                    masterData.PRVBE = worksheet.Cells[row, 74].GetStringOrNull();
                    masterData.NFEAG = worksheet.Cells[row, 75].GetStringOrNull();
                    masterData.NFGRP = worksheet.Cells[row, 76].GetStringOrNull();
                    masterData.KZKUP = worksheet.Cells[row, 77].GetStringOrNull();
                    masterData.INTRM = worksheet.Cells[row, 78].GetStringOrNull();
                    masterData.KZCLB = worksheet.Cells[row, 79].GetStringOrNull();
                    masterData.NLFZV = worksheet.Cells[row, 80].GetStringOrNull();
                    masterData.NLFMV = worksheet.Cells[row, 81].GetStringOrNull();
                    masterData.RFPNT = worksheet.Cells[row, 82].GetStringOrNull();
                    masterData.NLFZT = worksheet.Cells[row, 83].GetStringOrNull();
                    masterData.ITSOB = worksheet.Cells[row, 84].GetStringOrNull();

                    // Add the mapped MasterDataModel object to the list
                    masterDataList.Add(masterData);
                }
            }

            return masterDataList;
        }

        /// <summary>
        /// Read Machine condition tab in excel to object
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="machineConfig"></param>
        /// <returns></returns>
        public List<MachineConditionModel> ReadMachineCondition(string filePath, TemplateConfigModel machineConfig)
        {
            List<MachineConditionModel> dataList = new List<MachineConditionModel>();

            FileInfo file = new FileInfo(filePath);
            using (ExcelPackage package = new ExcelPackage(file))
            {
                int worksheetIndex = ExcelHelper.FindWorksheetIndex(package, machineConfig.TabName);
                ExcelWorksheet worksheet = package.Workbook.Worksheets[worksheetIndex];
                // calculate formula in worksheet before extract data
                worksheet.Calculate();
                int rowCount = machineConfig.EndLine;

                for (int row = machineConfig.StartLine; row <= rowCount; row++) // Start from row StartLine to skip headers
                {
                    MachineConditionModel data = new MachineConditionModel();

                    data.Line = worksheet.Cells[row, 1].GetStringOrNull();
                    data.Template = worksheet.Cells[row, 3].GetStringOrNull();
                    data.TemplateNLine = worksheet.Cells[row, 4].GetStringOrNull();
                    data.Column0230N0410 = worksheet.Cells[row, 6].GetStringOrNull();
                    data.Column0240N0490N0420N0310 = worksheet.Cells[row, 8].GetStringOrNull();
                    data.Column0260N0440 = worksheet.Cells[row, 9].GetStringOrNull();
                    data.Column0280N0460 = worksheet.Cells[row, 11].GetStringOrNull();
                    data.Column0290N0470 = worksheet.Cells[row, 13].GetStringOrNull();
                    data.Column0590N720 = worksheet.Cells[row, 15].GetStringOrNull();
                    data.Column0600N0730 = worksheet.Cells[row, 17].GetStringOrNull();
                    data.Column0610N0740 = worksheet.Cells[row, 19].GetStringOrNull();
                    data.Column0650N0780 = worksheet.Cells[row, 21].GetStringOrNull();
                    data.Column0220N0400 = worksheet.Cells[row, 23].GetStringOrNull();

                    dataList.Add(data);
                }
            }

            return dataList;
        }

        /// <summary>
        /// Read template file to object
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="templateConfig"></param>
        /// <param name="materialType"></param>
        /// <returns></returns>
        public List<TemplateModel> ReadTemplates(string filePath, TemplateConfigModel templateConfig, string materialType)
        {
            List<TemplateModel> templatess = new List<TemplateModel>();

            using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
            {
                int worksheetIndex = ExcelHelper.FindWorksheetIndex(package, templateConfig.TabName);
                ExcelWorksheet worksheet = package.Workbook.Worksheets[worksheetIndex];

                for (int row = templateConfig.StartLine; row <= templateConfig.EndLine; row++) // Assuming data starts from the second row
                {
                    int startResultColumn = 18;
                    startResultColumn = materialType == MaterialType.VPowder ? startResultColumn : startResultColumn + 1;
                    TemplateModel vPowderTemplate = new TemplateModel();
                    try
                    {
                        vPowderTemplate.MaterialNo = Convert.ToInt32(worksheet.Cells[row, 1].Value);
                        vPowderTemplate.Code = worksheet.Cells[row, 2].GetStringOrNull();
                        vPowderTemplate.Results1 = worksheet.Cells[row, startResultColumn++].GetStringOrNull();
                        vPowderTemplate.Results2 = worksheet.Cells[row, startResultColumn++].GetStringOrNull();
                        vPowderTemplate.Results3 = worksheet.Cells[row, startResultColumn++].GetStringOrNull();
                        vPowderTemplate.Results4 = worksheet.Cells[row, startResultColumn++].GetStringOrNull();
                        vPowderTemplate.Results5 = worksheet.Cells[row, startResultColumn++].GetStringOrNull();
                        vPowderTemplate.Results6 = worksheet.Cells[row, startResultColumn++].GetStringOrNull();
                        vPowderTemplate.Results7 = worksheet.Cells[row, startResultColumn++].GetStringOrNull();
                        vPowderTemplate.Results8 = worksheet.Cells[row, startResultColumn].GetStringOrNull();

                        templatess.Add(vPowderTemplate);
                    }
                    catch (Exception ex)
                    {
                        // log here
                        LogWarning(ex.Message + string.Format(materialType + " " + LogMessage.InvalidMaterialNo, worksheet.Cells[row, 1].Value, row));
                    }
                }
            }

            return templatess;
        }


        public void WriteOutputFile(string filePath, List<MasterDataModel> materialFinalized)
        {

            using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int startRow = 8;
                foreach (MasterDataModel data in materialFinalized)
                {
                    // format quantity
                    worksheet.Cells[startRow, 24].Style.Numberformat.Format = "0.000";
                    if (startRow > 1048576)
                    {
                        // find how to handle new tab
                        LogError(LogMessage.ReachMaximumRowSupport);
                        break;
                    }
                    int col = 1;
                    foreach (PropertyInfo property in typeof(MasterDataModel).GetProperties())
                    {
                        worksheet.Cells[startRow, col].Value = property.GetValue(data);
                        col++;
                    }
                    startRow++;
                }

                // Save the Excel file
                FileInfo file = new FileInfo(filePath);
                package.SaveAs(file);
            }
        }

        public void Process(string bomFilePath, string masterFilePath, string masterRecipeFilePath, VBAConfigModel config)
        {
            try
            {
                LogInfo(LogMessage.ProcessStart);
                // read master template
                List<MasterDataModel> masterTemplate = ReadMasterTemplate(masterFilePath, config.MasterTemplateConfig);
                // read machine condition
                List<MachineConditionModel> machineCondition = ReadMachineCondition(masterRecipeFilePath, config.MachineConditionConfig);
                // read vpowder template
                List<TemplateModel> vPowderTemplate = ReadTemplates(masterRecipeFilePath, config.VPowderTemplateConfig, MaterialType.VPowder);
                List<TemplateModel> interponTemplate = ReadTemplates(masterRecipeFilePath, config.InterponTemplateConfig, MaterialType.Interpon);
                // extract raw material
                List<MaterialModel> rawMaterial = ExtractRawMaterial(bomFilePath, config.BomFileConfig.RawTabName,
                    config.BomFileConfig.RawTabStartLine);

                // cut only selected bom
                BomModel bom = new BomModel();
                // extract to VPOWDER object with material scope 
                foreach (ExtractScopeRange vPowderScopeRange in config.BomFileConfig.VPowderScopeRanges)
                {
                    bom.VPowderModels.AddRange(ExtractMaterialRange(rawMaterial, vPowderScopeRange.StartRange,
                        vPowderScopeRange.EndRange));
                }
                // extract to INTERPON object with material scope
                foreach (ExtractScopeRange interponScopeRange in config.BomFileConfig.InterponScopeRanges)
                {
                    bom.InterPonModels.AddRange(ExtractMaterialRange(rawMaterial, interponScopeRange.StartRange,
                        interponScopeRange.EndRange));
                }

                // create new output file from template
                string filePath = FileHelper.GetOutputXlsxFile();

                IEnumerable<IGrouping<int, MaterialModel>> groupedVPowder = bom.VPowderModels.GroupBy(m => m.Material);
                IEnumerable<IGrouping<int, MaterialModel>> groupedInterpon = bom.InterPonModels.GroupBy(m => m.Material);
                List<MasterDataModel> materialFinalized = new List<MasterDataModel>();
                // convert using business logic for vpowder
                var vPowderResult = Cooking(masterTemplate, groupedVPowder.ToList(), machineCondition, vPowderTemplate, MaterialType.VPowder);
                // add vpowder to result
                materialFinalized.AddRange(vPowderResult);
                // convert using business logic for interpon
                var interponResult = Cooking(masterTemplate, groupedInterpon.ToList(), machineCondition, interponTemplate,
                    MaterialType.Interpon);
                // add intgerpon to result
                materialFinalized.AddRange(interponResult);
                CleanupData(materialFinalized);
                LogInfo(string.Format(LogMessage.ExcelWritingStart, filePath));
                // write xlsx file from object
                WriteOutputFile(filePath, materialFinalized);
                LogInfo(LogMessage.ExcelWritingFinished);
                LogInfo("Missing Material Template can be found at : " + _logger.LogPath);
                LogInfo(LogMessage.ProcessCompleted);
            }
            catch (Exception ex)
            {
                LogError(ex.Message);
            }
        }

        /// <summary>
        /// This method will clean up unused data, record and all unneeded data before write to excel file.
        /// </summary>
        /// <param name="materialFinalized"></param>
        private void CleanupData(List<MasterDataModel> materialFinalized)
        {
            var groupByMaterialNumber = materialFinalized.GroupBy(x => x.MATNR);
            List<string> materialRefreshScope = new List<string>();
            foreach (IGrouping<string, MasterDataModel> material in groupByMaterialNumber)
            {
                var groupByBom = material.GroupBy(x => x.STLAL);
                foreach (IGrouping<string, MasterDataModel> setOfAltBom in groupByBom)
                {
                    foreach (MasterDataModel masterDataModel in setOfAltBom.ToList())
                    {
                        if (_updateComponents.Contains(masterDataModel.POSNR) && masterDataModel.MENGE.HasValue == false)
                        {
                            materialRefreshScope.Add(masterDataModel.MATNR);
                            materialFinalized.Remove(masterDataModel);
                        }
                    }
                }
            }

            materialRefreshScope = materialRefreshScope.Distinct().ToList();
            // rerun pos number
            RefreshPosNumber(groupByMaterialNumber, materialRefreshScope);
        }

        /// <summary>
        /// Rerun POS number for all material by bom
        /// </summary>
        /// <param name="materialFinalizedGrouped"></param>
        /// <param name="materialRefreshScope"></param>
        private void RefreshPosNumber(IEnumerable<IGrouping<string, MasterDataModel>> materialFinalizedGrouped, List<string> materialRefreshScope)
        {
            foreach (IGrouping<string, MasterDataModel> material in materialFinalizedGrouped)
            {
                // select only relate material key
                if (materialRefreshScope.Contains(material.Key))
                {
                    // group material by bom and looping
                    var groupByBom = material.GroupBy(x => x.STLAL);
                    foreach (IGrouping<string, MasterDataModel> setOfAltBom in groupByBom)
                    {
                        // pos no start at 10
                        int posNumber = 10;
                        foreach (var bomMaterial in setOfAltBom.ToList())
                        {
                            // format to string 4 digits
                            string lineNo = posNumber.ToString("D4");
                            bomMaterial.POSNR = lineNo;
                            bomMaterial.SORTF = lineNo;
                            // increase pos no by 10
                            posNumber += 10;
                        }
                    }
                }
            }
        }


        /// <summary>
        /// Mapping method with all business logic to convert data by multiple source
        /// </summary>
        /// <returns></returns>
        public List<MasterDataModel> Cooking(List<MasterDataModel> masterData, List<IGrouping<int, MaterialModel>> bomMaterialMaster, List<MachineConditionModel> machineConditions, List<TemplateModel> templates, string materialType)
        {
            // add convert data here
            List<MasterDataModel> result = new List<MasterDataModel>();
            int materialSuccessCount = 0;
            foreach (IGrouping<int, MaterialModel> materialMaster in bomMaterialMaster)
            {
                var template = templates.FirstOrDefault(x => x.MaterialNo == materialMaster.Key);
                LogInfo(string.Format(LogMessage.ProcessingMaterialNoStart, materialMaster.Key));
                if (template == null) // template not found case
                {
                    // log warning
                    LogWarning($"{LogType.InvalidData}" + string.Format(materialType + " " + LogMessage.TemplateNotFoundOnMaterial, materialMaster.Key));
                    _logger.LogMissingMaterialNo(materialMaster.Key.ToString());
                    continue;
                }

                var groupedRawMaterial = materialMaster.ToList();
                // setup data for each template
                if (template.Results1 != null)
                {
                    MachineConditionModel machineCondition = GetMachineCondition(machineConditions, template.Results1);
                    ValidateMachineCondition(machineCondition, template.Results1);
                    
                    PrepareData(2, machineCondition, materialMaster.Key, masterData, groupedRawMaterial, result);
                }

                if (template.Results2 != null)
                {
                    MachineConditionModel machineCondition = GetMachineCondition(machineConditions, template.Results2);
                    ValidateMachineCondition(machineCondition, template.Results2);
                    PrepareData(3, machineCondition, materialMaster.Key, masterData, groupedRawMaterial, result);
                }

                if (template.Results3 != null)
                {
                    MachineConditionModel machineCondition = GetMachineCondition(machineConditions, template.Results3);
                    ValidateMachineCondition(machineCondition, template.Results3);
                    PrepareData(4, machineCondition, materialMaster.Key, masterData, groupedRawMaterial, result);
                }

                if (template.Results4 != null)
                {
                    MachineConditionModel machineCondition = GetMachineCondition(machineConditions, template.Results4);
                    ValidateMachineCondition(machineCondition, template.Results4);
                    PrepareData(5, machineCondition, materialMaster.Key, masterData, groupedRawMaterial, result);
                }

                if (template.Results5 != null)
                {
                    MachineConditionModel machineCondition = GetMachineCondition(machineConditions, template.Results5);
                    ValidateMachineCondition(machineCondition, template.Results5);
                    PrepareData(6, machineCondition, materialMaster.Key, masterData, groupedRawMaterial, result);
                }

                if (template.Results6 != null)
                {
                    MachineConditionModel machineCondition = GetMachineCondition(machineConditions, template.Results6);
                    ValidateMachineCondition(machineCondition, template.Results6);
                    PrepareData(7, machineCondition, materialMaster.Key, masterData, groupedRawMaterial, result);
                }

                if (template.Results7 != null)
                {
                    MachineConditionModel machineCondition = GetMachineCondition(machineConditions, template.Results7);
                    ValidateMachineCondition(machineCondition, template.Results7);
                    PrepareData(8, machineCondition, materialMaster.Key, masterData, groupedRawMaterial, result);
                }

                if (template.Results8 != null)
                {
                    MachineConditionModel machineCondition = GetMachineCondition(machineConditions, template.Results8);
                    ValidateMachineCondition(machineCondition, template.Results8);
                    PrepareData(9, machineCondition, materialMaster.Key, masterData, groupedRawMaterial, result);
                }

                LogSuccess(string.Format(LogMessage.ProcessingMaterialNoFinish, materialMaster.Key));
                materialSuccessCount+=1;
            }

            int totalVPowder = bomMaterialMaster.Count();
            LogSuccess(string.Format(materialType + " " + LogMessage.VPowderResult, totalVPowder, materialSuccessCount, totalVPowder - materialSuccessCount));
            return result;
        }

        private void ValidateMachineCondition(MachineConditionModel machineCondition, string templateName)
        {
            if (machineCondition == null)
            {
                // log warning
                LogWarning(string.Format(LogMessage.MachineConditionNotFound, templateName));
                _logger.LogMissingMachineCondition(templateName);
            }
        }

        private MachineConditionModel GetMachineCondition(List<MachineConditionModel> machineConditions, string templateName)
        {
            templateName = Regex.Replace(templateName, @"\s+", "");
            var template = machineConditions.FirstOrDefault(x => x.Template == templateName);
            if (template == null)
            {
                // log template not found
                LogWarning($"{LogType.InvalidData}" + String.Format(LogMessage.TemplateNotFound, templateName));
            }

            return template;
        }

        /// <summary>
        /// Mapping main component include material depend on each pos no.
        /// </summary>
        /// <param name="bomItemNumber"></param>
        /// <param name="vPowderModel"></param>
        private void UpdateComponents(MasterDataModel bomItemNumber, List<MaterialModel> vPowderModel)
        {
            MaterialModel mapMaterial = null;

            switch (bomItemNumber.POSNR)
            {
                case ComponentNumbers.Number0040:
                    mapMaterial = vPowderModel.FirstOrDefault(x => x.Item == ComponentNumbers.Number0010);
                    break;
                case ComponentNumbers.Number0050:
                    mapMaterial = vPowderModel.FirstOrDefault(x => x.Item == ComponentNumbers.Number0020);
                    break;
                case ComponentNumbers.Number0060:
                    mapMaterial = vPowderModel.FirstOrDefault(x => x.Item == ComponentNumbers.Number0030);
                    break;
                case ComponentNumbers.Number0070:
                    mapMaterial = vPowderModel.FirstOrDefault(x => x.Item == ComponentNumbers.Number0040);
                    break;
                case ComponentNumbers.Number0080:
                    mapMaterial = vPowderModel.FirstOrDefault(x => x.Item == ComponentNumbers.Number0050);
                    break;
                case ComponentNumbers.Number0090:
                    mapMaterial = vPowderModel.FirstOrDefault(x => x.Item == ComponentNumbers.Number0060);
                    break;
                case ComponentNumbers.Number0100:
                    mapMaterial = vPowderModel.FirstOrDefault(x => x.Item == ComponentNumbers.Number0070);
                    break;
                case ComponentNumbers.Number0110:
                    mapMaterial = vPowderModel.FirstOrDefault(x => x.Item == ComponentNumbers.Number0080);
                    break;
            }
            if (mapMaterial != null)
            {
                // mapping data
                bomItemNumber.IDNRK = mapMaterial.Component;
                bomItemNumber.MEINS = mapMaterial.Un;
                bomItemNumber.MENGE = mapMaterial.Quantity1;
                bomItemNumber.POTX1 = mapMaterial.BOMComponent;
            }
            else
            {
                LogWarning($"{LogType.InvalidData}" + string.Format(LogMessage.TemplateNotFoundOnMaterial, bomItemNumber.MATNR + $" POSNR : {bomItemNumber.POSNR}"));
            }
        }

        /// <summary>
        /// This function contains the main mapping data
        /// </summary>
        /// <param name="alternateBomNo"></param>
        /// <param name="machineCondition"></param>
        /// <param name="materialNo"></param>
        /// <param name="masterData"></param>
        /// <param name="vPowderRawGroup"></param>
        /// <param name="result"></param>
        private void PrepareData(int alternateBomNo, MachineConditionModel machineCondition, int materialNo, List<MasterDataModel> masterData, List<MaterialModel> vPowderRawGroup, List<MasterDataModel> result)
        {
            var alternateBom = ClassExtensions.DeepCopy(masterData);
            foreach (var item in alternateBom)
            {
                // mapping data
                item.STLAL = alternateBomNo.ToString();
                item.MATNR = materialNo.ToString();
                if (_updateComponents.Contains(item.POSNR))
                {
                    UpdateComponents(item, vPowderRawGroup);
                }
                if (machineCondition != null)
                {
                    UpdateOtherComponents(item, machineCondition);
                    // mapping data
                    item.ProductVersion = machineCondition.Line;
                    item.PLNNR = "UB1P-" + FormatProductNumber(item.ProductVersion);
                }
            }

            result.AddRange(alternateBom);
        }

        /// <summary>
        /// This function will map pos no following "Machine Condition" in 02 vba master recipe.xlsx file
        /// </summary>
        /// <param name="bomItemNumber"></param>
        /// <param name="machineCondition"></param>
        private void UpdateOtherComponents(MasterDataModel bomItemNumber, MachineConditionModel machineCondition)
        {
            string bomText = string.Empty;
            switch (bomItemNumber.POSNR)
            {
                case ComponentNumbers.Number0410:
                case ComponentNumbers.Number0230:
                    bomText = machineCondition.Column0230N0410;
                    break;
                case ComponentNumbers.Number0490:
                case ComponentNumbers.Number0420:
                case ComponentNumbers.Number0310:
                case ComponentNumbers.Number0240:
                    bomText = machineCondition.Column0240N0490N0420N0310;
                    break;
                case ComponentNumbers.Number0440:
                case ComponentNumbers.Number0260:
                    bomText = machineCondition.Column0260N0440;
                    break;
                case ComponentNumbers.Number0460:
                case ComponentNumbers.Number0280:
                    bomText = machineCondition.Column0280N0460;
                    break;
                case ComponentNumbers.Number0470:
                case ComponentNumbers.Number0290:
                    bomText = machineCondition.Column0290N0470;
                    break;
                case ComponentNumbers.Number0720:
                case ComponentNumbers.Number0590:
                    bomText = machineCondition.Column0590N720;
                    break;
                case ComponentNumbers.Number0730:
                case ComponentNumbers.Number0600:
                    bomText = machineCondition.Column0600N0730;
                    break;
                case ComponentNumbers.Number0740:
                case ComponentNumbers.Number0610:
                    bomText = machineCondition.Column0610N0740;
                    break;
                case ComponentNumbers.Number0780:
                case ComponentNumbers.Number0650:
                    bomText = machineCondition.Column0650N0780;
                    break;
                case ComponentNumbers.Number0400:
                case ComponentNumbers.Number0220:
                    bomText = machineCondition.Column0220N0400;
                    break;
            }

            bomItemNumber.POTX1 = bomText == string.Empty ? bomItemNumber.POTX1 : bomText;
        }

        private string FormatProductNumber(string input)
        {
            string pattern = @"(\D+)(0*)(\d+)";

            // Replace the numeric part with leading zeros if it has less than two digits
            string result = Regex.Replace(input, pattern, match =>
            {
                int number = int.Parse(match.Groups[3].Value);
                return $"{match.Groups[1].Value}{number:D2}";
            });

            return result;
        }

        public void LogSuccess(string message) => Log(EnumLogLevel.Success, message);
        public void LogInfo(string message) => Log(EnumLogLevel.Info, message);
        public void LogWarning(string message) => Log(EnumLogLevel.Warning, message);
        public void LogError(string message) => Log(EnumLogLevel.Error, message);

        public void Log(EnumLogLevel level, string message)
        {
            // Raise the event with severity
            LogGenerated?.Invoke(this, Tuple.Create(message, level));
        }
    }
}
