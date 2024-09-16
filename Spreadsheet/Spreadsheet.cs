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
                if (file is not { Length: > 0 })
                    throw new Exception("File is corrupted");

                List<T> response = [];
                PropertyInfo[] properties = typeof(T).GetProperties();

                using var workbook = new XLWorkbook(file);
                var worksheet = workbook.Worksheet(sheetName);

                var header = worksheet.Row(headerRowIndex);

                foreach (var row in worksheet.RowsUsed().Skip(startRowIndex > 1 ? startRowIndex - 1 : headerRowIndex ))
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
                                if (attribute is not SpreadsheetFieldAttribute fieldAttribute ||
                                    fieldAttribute.CellName != headerField.GetString()) continue;

                                PropertyInfo tResponseInfo = typeof(T).GetProperty(property.Name);
                                if (tResponseInfo.CanWrite)
                                {
                                    TypeCode typeCode = Type.GetTypeCode(tResponseInfo.PropertyType);

                                    switch (typeCode)
                                    {
                                        case TypeCode.String:
                                            if (fieldAttribute.Length > 0 &&
                                                cell.Value.ToString().Length != fieldAttribute.Length)
                                                throw new Exception("Length is not match");

                                            if (!string.IsNullOrEmpty(fieldAttribute.StartWith) &&
                                                !cell.Value.ToString().StartsWith(fieldAttribute.StartWith))
                                                throw new Exception("StartWith is not match");

                                            tResponseInfo.SetValue(TResponse, cell.GetString());
                                            break;
                                        case TypeCode.Int32:
                                            tResponseInfo.SetValue(TResponse, (int)cell.Value);
                                            break;
                                        case TypeCode.Boolean:
                                            tResponseInfo.SetValue(TResponse, (bool)cell.Value);
                                            break;
                                        case TypeCode.Double:
                                            tResponseInfo.SetValue(TResponse, (double)cell.Value);
                                            break;
                                        case TypeCode.Int64:
                                            tResponseInfo.SetValue(TResponse, (long)cell.Value);
                                            break;
                                        case TypeCode.Decimal:
                                            tResponseInfo.SetValue(TResponse, decimal.Parse(cell.Value.ToString()));
                                            break;
                                        case TypeCode.DateTime:
                                            tResponseInfo.SetValue(TResponse, DateTime.Parse(cell.Value.ToString()));
                                            break;
                                        default:
                                            tResponseInfo.SetValue(TResponse, cell.GetString());
                                            break;
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

        public MemoryStream Write(List<T> data, string sheetName, int headerRowIndex = 0)
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
                            foreach (var item in data)
                            {
                                PropertyInfo tResponseInfo = typeof(T).GetProperty(property.Name);
                                worksheet.Cell(rowIndex, i).Value = tResponseInfo.GetValue(item)?.ToString();
                                rowIndex++;
                            }
                        }

                        i++;
                    }
                }
                
                string path = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "/excel.xlsx";
                workbook.SaveAs(path);
                
                MemoryStream file = new();
                workbook.SaveAs(file);
                return file;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        public void Fill(MemoryStream file, List<T> data, string sheetName, int headerRowIndex = 0,
            int dataRowIndex = 0)
        {
            try
            {
                if (file == null || file.Length <= 0)
                    throw new Exception("File is corrupted");

                if (data == null || data.Count == 0)
                    throw new Exception("Data is null");

                using var workbook = new XLWorkbook(file);
                var worksheet = workbook.Worksheet(sheetName);

                var header = worksheet.Row(headerRowIndex);

                PropertyInfo[] properties = typeof(T).GetProperties();

                foreach (var property in properties)
                {
                    object[] attributes = property.GetCustomAttributes(typeof(SpreadsheetFieldAttribute), true);
                    foreach (var attribute in attributes)
                    {
                        if (attribute is SpreadsheetFieldAttribute fieldAttribute)
                        {
                            var columnIndex = header.CellsUsed()
                                .FirstOrDefault(x => x.Value.ToString() == property.Name);
                            int rowIndex = dataRowIndex;
                            foreach (var item in data)
                            {
                                PropertyInfo tResponseInfo = typeof(T).GetProperty(property.Name);
                                worksheet.Cell(rowIndex, columnIndex.Address.ColumnNumber).Value =
                                    tResponseInfo.GetValue(item).ToString();
                                rowIndex++;
                            }
                        }
                    }
                }

                string path = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "/Dupexcel.xlsx";
                workbook.SaveAs(path);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
    }
}