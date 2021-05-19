using ExcelDataReader;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Hosting;
using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace BlazorFiles.Api.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExcelController : ControllerBase
    {
        private readonly IHostEnvironment _environment;

        public ExcelController(IHostEnvironment environment)
        {
            _environment = environment;
        }

        [HttpPost]
        public async Task<IActionResult> Post([FromForm] IFormFile excel)
        {
            if (excel == null || excel.Length == 0)
                return BadRequest("Upload a file");

            string fileName = excel.FileName;
            string extension = Path.GetExtension(fileName);

            string[] allowedExtensions = { ".xlsx", ".xls" };

            if (!allowedExtensions.Contains(extension.ToLower()))
                return BadRequest("File is not a valid excel");

            string newFileName = $"{Guid.NewGuid()}{extension}";
            string filePath = Path.Combine(_environment.ContentRootPath, "wwwroot", "Excel", newFileName);

            using (var fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                await excel.CopyToAsync(fileStream).ConfigureAwait(false);
                _ = fileStream.FlushAsync();
            }

            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            using (var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                var reader = ExcelReaderFactory.CreateReader(fileStream,
                new ExcelReaderConfiguration() { FallbackEncoding = System.Text.Encoding.GetEncoding(1252) });

                var dataset = reader.AsDataSet(new ExcelDataSetConfiguration()
                {
                    ConfigureDataTable = _ => new ExcelDataTableConfiguration()
                    {
                        //FilterRow = rowReader => rowReader != default,
                        FilterColumn = (rowReader, columnIndex) => rowReader[columnIndex] != default,
                        EmptyColumnNamePrefix = "Column",
                        UseHeaderRow = true,
                        ReadHeaderRow = (rowReader) =>
                        {
                            rowReader.Read();
                            rowReader.Read();
                            rowReader.Read();
                        },
                    }
                });

                foreach (DataTable table in dataset.Tables)
                {
                    var name = table.TableName;

                    int colCount = table.Columns.Count;
                    int rowCount = table.Rows.Count;
                    int count = 0;

                    for (int row = 0; row < rowCount; row++)
                    {
                        DataRow rowData = table.Rows[row];
                        var itemarray = rowData.ItemArray;
                        foreach (var item in itemarray)
                        {
                            if (item.ToString().Length == 0)
                            {
                                count++;
                            }
                        }
                        if (count == itemarray.Length)
                        {
                            rowData.Delete();
                        }
                        int colCount1 = rowData.Table.Columns.Count;
                        int rowCount1 = rowData.Table.Rows.Count;
                        if (count != itemarray.Length)
                        {
                            for (int col = 0; col < colCount1; col++)
                            {
                                DataColumn colData = table.Columns[col];
                                var rowVar = rowData[col];

                                Console.WriteLine(" Row:" + row + " column:" + col + " Value:" + colData + rowVar);
                                count = 0;
                            }
                        }
                        count = 0;
                    }
                }
            }

            return Ok($"Excel/{newFileName}");
        }
    }
}
