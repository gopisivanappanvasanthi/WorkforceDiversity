using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using System.Data;
using System.Text;
using System.Collections.Generic;
using ExcelDataReader;

namespace azurefuncconverter
{
    public static class XLtoJson
    {
        [FunctionName("XLtoJson")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");

            //Initialise variables
            DataSet result = new DataSet();
            string workSheetTab = req.Query["workSheetTab"];
            string filename = req.Query["filename"];
            string extension = System.IO.Path.GetExtension(filename.ToLower());
            int workSheetIndex = 0;
            string requestBody = await new StreamReader(req.Body).ReadToEndAsync();

            //ExcelDataReader throws a NotSupportedException "No data is available for encoding 1252." on .NET Core.
            //Unless add code to register the code page provider
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            if (requestBody != null && extension == ".xls")
            {
                log.LogInformation("Reading from a binary Excel file ('97-2003 format;*" + extension);
                IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(req.Body, null);
                excelReader.Read();
                excelReader.Close();
            }
            else if (requestBody != null && extension == ".xlsx")
            {
                log.LogInformation("Reading from a binary Excel file (format; *" + extension + ")");
                IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(req.Body, null);
                excelReader.Read();
                excelReader.Close();
            }
            else
            {
                log.LogInformation("Returning from a binary file (format; *" + extension);
                //TODO 
            }

            //List of worksheet tabs
            List<string> items = new List<string>();
            for (int i = 0; i < result.Tables.Count; i++)
                items.Add(result.Tables[i].TableName.ToString());

            //Try block required
            //Select Work Sheet Tab Index
            log.LogInformation("WorkSheetTab = " + workSheetTab);
            if (int.TryParse(workSheetTab, out workSheetIndex))
            {
            }
            else
            {
                workSheetIndex = items.IndexOf(workSheetTab);
            }

            //Convert one worksheet tab to CSV file
            log.LogInformation("START CONVERTTOCSV");
            string csv = convertToCSV(result, workSheetIndex);
            log.LogInformation("END CONVERTTOCSV");


            return workSheetTab != null && filename != null
               ? (ActionResult)new OkObjectResult(filename)
                : new BadRequestObjectResult("Please pass the workSheetTab to be converted & the filename on the query string or in the request body");

        }

        private static string convertToCSV(DataSet result, int ind)
        {
            // sheets in excel file becomes tables in dataset
            //result.Tables[0].TableName.ToString(); // to get sheet name (table name)

            int row_no = 0;
            StringBuilder builder = new StringBuilder();
            const string ap = ",";
            const string nl = "\r\n";
            while (row_no < result.Tables[ind].Rows.Count)
            {

                for (int i = 0; i < result.Tables[ind].Columns.Count - 1; i++)
                {
                    string val = result.Tables[ind].Rows[row_no][i].ToString();
                    if (i != 0)
                    {
                        builder.Append(ap);
                    }
                    if (val.Contains(ap))
                    {
                        builder.Append('"');
                        builder.Append(val);
                        builder.Append('"');

                    }
                    else
                    {
                        builder.Append(val);
                    }
                }
                row_no++;
                builder.Append(nl);
            }
            return builder.ToString();
        }
    }
}
