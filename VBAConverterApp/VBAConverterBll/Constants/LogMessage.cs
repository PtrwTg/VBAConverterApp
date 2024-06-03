namespace VBAConverterApp.VBAConverterBll.Constants
{
    public static class LogMessage
    {
        public const string TemplateNotFound = "Template Not Found :{0}";
        public const string TemplateNotFoundOnMaterial = "Template Not Found on Material Number : {0} ";
        public const string WorksheetTabNotFound = "Worksheet tab not found on tab name : {0} ";
        public const string InvalidMaterialNo = "Material cannot be convert : {0} with Row No. : {1} ";
        public const string ReachMaximumRowSupport = "Excel row reach maximum row at 1048576, please check";
        public const string MachineConditionNotFound = "Machine Condition template {0} was not found.";
        public const string ProcessStart = "Processing Started.";
        public const string ProcessCompleted = "Processing completed.";
        public const string ProcessingMaterialNoStart = "Processing Material No. : {0} .";
        public const string ProcessingMaterialNoFinish = "Processing Material No. : {0} Finished.";
        public const string ExcelWritingStart = "Writting out file at {0}.";
        public const string ExcelWritingFinished = "Writting excel file finished.";
        public const string VPowderResult = "Total items : {0}. Processed items :{1}. Error items : {2}.";
    }
}