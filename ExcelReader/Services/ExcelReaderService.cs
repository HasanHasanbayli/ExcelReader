using ClosedXML.Excel;
using Microsoft.AspNetCore.Http;

namespace ExcelReader.Services;

public class FileReaderService
{
    private Dictionary<string, Dictionary<string, object?>> Parameters { get; set; } = new();
    private IXLWorksheet Worksheet { get; set; } = default!;
    private string WorkSheetName { get; set; } = default!;

    public async Task<(Dictionary<string, (IEnumerable<string>, IEnumerable<IEnumerable<object>>)> worksheets, IEnumerable<string>errors)>
        ReadExcelFileAsync(IFormFile file, Dictionary<string, Dictionary<string, object?>> parameters)
    {
        Dictionary<string, (IEnumerable<string>, IEnumerable<IEnumerable<object>>)> worksheets = new();
        List<string> errors = [];
        Parameters = parameters;
        try
        {
            await using Stream stream = file.OpenReadStream();

            using XLWorkbook workbook = new(stream);

            foreach (string worksheetName in Parameters.Keys.ToList())
            {
                try
                {
                    Worksheet = workbook.Worksheet(worksheetName);

                    if (Worksheet is null)
                    {
                        errors.Add(item: $"Worksheet '{worksheetName}' not found in Excel file.");
                        continue;
                    }

                    WorkSheetName = worksheetName;

                    IEnumerable<string> headers = ReadHeaders();

                    IEnumerable<IEnumerable<object>> data = ReadData();

                    worksheets.Add(worksheetName, (headers, data));
                }
                catch (Exception ex)
                {
                    errors.Add(item: $"Error reading Excel file: {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            errors.Add(item: $"Error opening Excel file: {ex.Message}");
        }

        return (worksheets, errors);
    }

    private IEnumerable<string> ReadHeaders()
    {
        bool hasHeaders = GetBoolParameterOrDefault(Parameters, key: "hasHeaders", defaultValue: true);

        int startHeaderRow = GetIntParameterOrDefault(Parameters, key: "headerRow");

        IEnumerable<int> cols = GetColumns(Parameters);

        return hasHeaders
            ? cols.Select(column => Worksheet.Cell(startHeaderRow, column).Value.ToString()).ToList()
            : [];
    }

    private IEnumerable<IEnumerable<object>> ReadData()
    {
        int startBodyRow = GetIntParameterOrDefault(Parameters, key: "bodyRow");

        int endBodyRow = Worksheet.LastRowUsed().RowNumber() -
                         Worksheet.FirstRowUsed().RowNumber() - startBodyRow;

        IEnumerable<int> cols = GetColumns(Parameters);

        return Enumerable
            .Range(startBodyRow, endBodyRow)
            .Select(row => cols.Select(col => GetCellValue(Worksheet.Cell(row, col))).ToList())
            .Where(IsRowValid)
            .ToList();
    }

    private IEnumerable<int> GetColumns(Dictionary<string, Dictionary<string, object?>> parameters)
    {
        if (!parameters[WorkSheetName].TryGetValue("cols", out object? obj))
        {
            int firstColumn = Worksheet.FirstColumnUsed().ColumnNumber();
            int lastColumn = Worksheet.LastColumnUsed().ColumnNumber();

            return Enumerable.Range(firstColumn, lastColumn - firstColumn + 1);
        }

        return obj switch
        {
            IEnumerable<int> numRange => numRange,
            string textRange => ParseColumnRange(textRange),
            _ => []
        };
    }

    private IEnumerable<int> ParseColumnRange(string columnRange)
    {
        string[] range = columnRange.Split(separator: ':');

        if (range.Length != 2)
            throw new ArgumentException(message: "Invalid column range.");

        int startCol = ExcelColumnToNumber(range[0]);
        int endCol = ExcelColumnToNumber(range[1]);

        return Enumerable.Range(startCol, endCol - startCol + 1);
    }

    private int ExcelColumnToNumber(string column)
    {
        int columnNumber = default;

        foreach (char c in column)
        {
            columnNumber *= 26;
            columnNumber += c - 'A' + 1;
        }

        return columnNumber;
    }

    private bool GetBoolParameterOrDefault(Dictionary<string, Dictionary<string, object?>> parameters, string key,
        bool defaultValue = false)
    {
        return !parameters[WorkSheetName].TryGetValue(key, out object? obj) || obj is not bool boolValue
            ? defaultValue
            : boolValue;
    }

    private int GetIntParameterOrDefault(Dictionary<string, Dictionary<string, object?>> parameters, string key,
        int defaultValue = default)
    {
        return !parameters[WorkSheetName].TryGetValue(key, out object? obj) || obj is not int intValue
            ? defaultValue
            : intValue;
    }

    private bool IsRowValid(IEnumerable<object> rowData)
    {
        bool rowIsExistAndHasValue = rowData != null && rowData.Any(data => data != default);

        return rowIsExistAndHasValue;
    }

    private object GetCellValue(IXLCell cell)
    {
        if (cell == default) return null!;

        return (cell.DataType switch
        {
            XLDataType.Blank => null,
            XLDataType.Boolean => cell.GetBoolean(),
            XLDataType.DateTime => cell.GetDateTime(),
            XLDataType.TimeSpan => cell.GetTimeSpan(),
            XLDataType.Number => cell.GetDouble(),
            XLDataType.Text => cell.GetString(),
            _ => null
        })!;
    }
}