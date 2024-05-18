using PackageRepository.Components.Spreadsheet;

namespace ExcelWorkshop
{
    public class WeatherForecast
    {
        [SpreadsheetField(cellName: "TemperatureC")]
        public int TemperatureC { get; set; }

        [SpreadsheetField(cellName: "Summary")]
        public string Summary { get; set; }
    }
}