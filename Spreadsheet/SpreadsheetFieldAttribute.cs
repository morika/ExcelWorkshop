namespace PackageRepository.Components.Spreadsheet
{
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Struct, AllowMultiple = false)]
    public class SpreadsheetFieldAttribute(string cellName, int length = 0, string startWith = "") : Attribute
    {
        public string CellName { get; } = cellName;
        public int Length { get; set; } = length;
        public string StartWith { get; set; } = startWith;
    }
}