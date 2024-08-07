using System.Reflection;
using ClosedXML.Excel;
using Microsoft.AspNetCore.Http.HttpResults;

namespace PackageRepository.Components.Spreadsheet
{
    public class Spreadsheet<T> where T : class, new()
    {
        public List<T> Read(MemoryStream file, string sheetName, int headerRowIndex = 1, int startRowIndex = 2)
        {
            try
            {
                if (file == null || file.Length <= 0)
                    throw new Exception("File is corrupted");

                List<T> response = [];
                PropertyInfo[] properties = typeof(T).GetProperties();

                using var workbook = new XLWorkbook(file);
                var worksheet = workbook.Worksheet(sheetName);

                var header = worksheet.Row(headerRowIndex);

                foreach (var row in worksheet.RowsUsed().Skip(startRowIndex))
                {
                    T TResponse = new();
                    foreach (var cell in row.CellsUsed())
                    {
                        var headerField = header.Cell(cell.Address.ColumnNumber);
                        foreach (var property in properties)
                        {
                            object[] attributes = property.GetCustomAttributes(typeof(SpreadsheetFieldAttribute), true);
                            foreach (var attribute in attributes)
                            {
                                if (attribute is SpreadsheetFieldAttribute fieldAttribute && fieldAttribute.CellName == headerField.GetString())
                                {
                                    PropertyInfo tResponseInfo = typeof(T).GetProperty(property.Name);
                                    if (tResponseInfo.CanWrite)
                                    {
                                        TypeCode typeCode = Type.GetTypeCode(tResponseInfo.PropertyType);
                                        switch (typeCode)
                                        {
                                            case TypeCode.Int32:
                                                tResponseInfo.SetValue(TResponse, (int)cell.Value);
                                                break;
                                            case TypeCode.Boolean:
                                                tResponseInfo.SetValue(TResponse, (bool)cell.Value);
                                                break;
                                            case TypeCode.Double:
                                                tResponseInfo.SetValue(TResponse, (double)cell.Value);
                                                break;
                                            default:
                                                tResponseInfo.SetValue(TResponse, cell.GetString());
                                                break;
                                        }

                                    }
                                }
                            }
                        }
                    }

                    response.Add(TResponse);
                }

                return response;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        public void Write(List<T> data, string sheetName, int headerRowIndex = 0)
        {
            try
            {
                if (data == null || data.Count == 0)
                    throw new Exception("Data is null");

                using var workbook = new XLWorkbook();
                var worksheet = workbook.Worksheets.Add(sheetName);

                PropertyInfo[] properties = typeof(T).GetProperties();
                int i = 1;

                foreach (var property in properties)
                {
                    object[] attributes = property.GetCustomAttributes(typeof(SpreadsheetFieldAttribute), true);
                    foreach (var attribute in attributes)
                    {
                        if (attribute is SpreadsheetFieldAttribute fieldAttribute)
                        {
                            worksheet.Cell(headerRowIndex, i).Value = fieldAttribute.CellName;
                            int rowIndex = headerRowIndex + 1;
                            foreach(var item in data)
                            {
                                PropertyInfo tResponseInfo = typeof(T).GetProperty(property.Name);
                                worksheet.Cell(rowIndex, i).Value = tResponseInfo.GetValue(item).ToString();
                                rowIndex++;
                            }
                        }

                        i++;
                    }
                }

                string path = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "/excel.xlsx";
                workbook.SaveAs(path);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
    }
}