using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using System.Data.SqlClient;
using Microsoft.Azure.WebJobs.Extensions.Storage;
using System.Data;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;

namespace RetrievingExcelInAzureFunc
{
    public static class Function1
    {
        [FunctionName("Function1")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequest req, 
             ILogger log)
        
        {
            log.LogInformation("C# HTTP trigger function processed a request.");

            // Get the directory path where your Excel files are stored
            //string file = @"C:\ExcelDoc\excelfile.xlsx";
            try
            {
                var formCollection = await req.ReadFormAsync();
                var file = formCollection.Files["file"];

                if (file == null || file.Length == 0)
                {
                    return new BadRequestObjectResult("Please upload a valid Excel file.");
                }

                // Assuming the file is uploaded to a temporary directory
                var filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".xlsx");

                using (var fileStream = new FileStream(filePath, FileMode.Create))
                {
                    await file.CopyToAsync(fileStream);
                }
                // Process the Excel file
                DataTable dataTable = ReadExcelData(filePath);

                // Do something with the DataTable, such as storing it in a database
                // StoreDataInDatabase(dataTable);

                //log.LogInformation($"Data from {file.FileName} processed successfully.");

                return new OkObjectResult("Excel file processed successfully.");
            }
            catch (Exception ex)
            {
                log.LogError($"An error occurred: {ex.Message}");
                return new StatusCodeResult(500);
            }
            static DataTable ReadExcelData(string filePath)
            {
                DataTable dataTable = new DataTable();

                using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
                {
                    WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                    WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                    SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                    foreach (Row row in sheetData.Elements<Row>())
                    {
                        if (row.RowIndex == 1) // Assuming the first row contains column headers
                        {
                            foreach (Cell cell in row.Elements<Cell>())
                            {
                                dataTable.Columns.Add(GetCellValue(workbookPart, cell));
                            }
                        }
                        else
                        {
                            DataRow dataRow = dataTable.NewRow();
                            int columnIndex = 0;

                            foreach (Cell cell in row.Elements<Cell>())
                            {
                                dataRow[columnIndex++] = GetCellValue(workbookPart, cell);
                            }

                            dataTable.Rows.Add(dataRow);
                        }
                    }
                }

                return dataTable;
            }

            static string GetCellValue(WorkbookPart workbookPart, Cell cell)
            {
                SharedStringTablePart sharedStringTablePart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
                string cellValue = cell.InnerText;

                if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
                {
                    int sharedStringIndex;
                    if (int.TryParse(cellValue, out sharedStringIndex))
                    {
                        cellValue = sharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(sharedStringIndex).InnerText;
                    }
                }

                return cellValue;
            }
        }



        
    }
}
