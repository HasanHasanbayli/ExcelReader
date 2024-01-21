# ExcelReader

ExcelReader is a .NET 8.0 project that provides services for reading Excel files. It uses the ClosedXML library to interact with Excel files and Microsoft.AspNetCore.Http for handling file uploads.

## Features

- Read Excel files asynchronously
- Extract data from specified columns
- Handle Excel files with or without headers
- Customizable parameters for reading data

## Usage

The main service provided by this project is `ExcelReaderService`. It has a method `ReadExcelFileAsync` which takes an `IFormFile` and a `Dictionary<string, Dictionary<string, object?>>` as parameters.

The `IFormFile` parameter is the Excel file you want to read.

The `Dictionary<string, Dictionary<string, object?>>` parameter is a dictionary of parameters for each worksheet in the Excel file. The key is the worksheet name and the value is another dictionary of parameters for that worksheet.

Here are the parameters you can specify for each worksheet:

- `hasHeaders`: A boolean indicating whether the worksheet has headers. Default is `true`.
- `headerRow`: The row number where the headers start. Default is `0`.
- `bodyRow`: The row number where the body data starts. Default is `0`.
- `cols`: The columns to read. Can be a range in the format "A:Z" or an `IEnumerable<int>` of column numbers.

## Dependencies

- ClosedXML 0.102.2
- Microsoft.AspNetCore.Http 2.2.2

## Setup

To use this project, you need to have .NET 8.0 installed on your machine. After cloning the repository, you can open the project in JetBrains Rider or any other .NET compatible IDE.

## License

This project is licensed under the terms of the MIT license.


To use the `ExcelReaderService`, you need to create an instance of it and call the `ReadExcelFileAsync` method. Here's a sample usage:

```csharp
using ExcelReader.Services;
using Microsoft.AspNetCore.Http;
using System.Collections.Generic;
using System.Threading.Tasks;

public class Program
{
    public static async Task Main()
    {
        // Create an instance of ExcelReaderService
        var excelReaderService = new ExcelReaderService();

        // Create a mock IFormFile (representing the Excel file)
        IFormFile file = new FormFile(new MemoryStream(File.ReadAllBytes("path_to_your_excel_file")), 0, 0, "Data", "filename.xlsx");

        // Define parameters for each worksheet
        var parameters = new Dictionary<string, Dictionary<string, object?>>
        {
            {
                "Worksheet1", new Dictionary<string, object?>
                {
                    { "hasHeaders", true },
                    { "headerRow", 0 },
                    { "bodyRow", 1 },
                    { "cols", "A:Z" }
                }
            },
            {
                "Worksheet2", new Dictionary<string, object?>
                {
                    { "hasHeaders", false },
                    { "headerRow", 0 },
                    { "bodyRow", 0 },
                    { "cols", new List<int> { 1, 2, 3 } }
                }
            }
        };

        // Call the ReadExcelFileAsync method
        var (worksheets, errors) = await excelReaderService.ReadExcelFileAsync(file, parameters);

        // Handle the results
        foreach (var worksheet in worksheets)
        {
            Console.WriteLine($"Worksheet: {worksheet.Key}");
            Console.WriteLine("Headers:");
            foreach (var header in worksheet.Value.Item1)
            {
                Console.WriteLine(header);
            }
            Console.WriteLine("Data:");
            foreach (var row in worksheet.Value.Item2)
            {
                foreach (var cell in row)
                {
                    Console.Write($"{cell} ");
                }
                Console.WriteLine();
            }
        }

        if (errors.Count > 0)
        {
            Console.WriteLine("Errors:");
            foreach (var error in errors)
            {
                Console.WriteLine(error);
            }
        }
    }
}
```

Please replace `"path_to_your_excel_file"` with the actual path to your Excel file. This example assumes that you have two worksheets named "Worksheet1" and "Worksheet2". Adjust the parameters according to your actual Excel file structure.