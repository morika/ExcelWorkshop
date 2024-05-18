namespace PackageRepository.Components.Spreadsheet
{
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Struct, AllowMultiple = false)]
    public class SpreadsheetFieldAttribute(string cellName) : Attribute
    {
        public string CellName { get; } = cellName;
    }
}