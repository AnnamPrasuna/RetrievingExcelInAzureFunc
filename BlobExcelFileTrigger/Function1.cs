using System;
using System.IO;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;

namespace BlobExcelFileTrigger
{
    public class Function1
    {
        [FunctionName("Function1")]
        public void Run([BlobTrigger("%ContainerName%", Connection = "AzureWebJobsStorage")]Stream myBlob, string name, ILogger log)
        {
            log.LogInformation($"C# Blob trigger function Processed blob\n Name:{name} \n Size: {myBlob.Length} Bytes");
            //get the file extension
            string extension = Path.GetExtension(name);
            //check if the file is in excel file
            if(extension==".xls"||extension==".xlsx")
            {
                //process the excel files
                using (var package = new ExcelPackage(myBlob))
                {
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    var worksheet = package.Workbook.Worksheets[0];
                    var rowcount = worksheet.Dimension.Columns;
                    for(int row=2;row<=rowcount;row++)
                    {
                        var phonenumber = worksheet.Cells[row, 1].Value?.ToString();
                        if(phonenumber!=null && phonenumber.Length==10)
                        {
                            var firstname = worksheet.Cells[row, 2].Value?.ToString();
                            var lastname = worksheet.Cells[row, 3].Value?.ToString();
                            var address = worksheet.Cells[row, 4].Value?.ToString();
                            var groupname = worksheet.Cells[row, 5].Value?.ToString();
                            var model = new MyModel()
                            {
                                Phonenumber = phonenumber,
                                Firstname = firstname,
                                Lastname = lastname,
                                Address = address,
                                Groupname = groupname,
                            };
                            log.LogInformation($"processed row{row-1}:{model}");
                        }
                    }
                }
            }
            else
            {
                log.LogInformation($"Ignoring blob {name} because it is not in an excel file");
            }
        }
    }
    public class MyModel
    {
        public string Phonenumber {  get; set; }
        public string Firstname { get; set; }
        public string Lastname { get; set; }
        public string Address { get; set; }
        public string Groupname { get; set; }
        public override string ToString()
        {
            return $"{{Phonenumber={Phonenumber},Firstname={Firstname},LastName={Lastname},Address={Address},GroupName={Groupname}}}";
        }
    }
}
