using PackageRepository.Components.Spreadsheet;

namespace ExcelWorkshop
{
    public class WeatherForecast
    {
        [SpreadsheetField(cellName: "TemperatureC")]
        public decimal TemperatureC { get; set; }

        [SpreadsheetField(cellName: "Summary", length: 10, startWith: "09")]
        public string Summary { get; set; }
    }
}