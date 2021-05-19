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

                //var dataset = reader.AsDataSet(new ExcelDataSetConfiguration()
                //{
                //    ConfigureDataTable = _ => new ExcelDataTableConfiguration()
                //    {
                //        //FilterRow = rowReader => rowReader != default,
                //        FilterColumn = (rowReader, columnIndex) => rowReader[columnIndex] != default,
                //        EmptyColumnNamePrefix = "Column",
                //        UseHeaderRow = true,
                //        ReadHeaderRow = (rowReader) =>
                //        {
                //            rowReader.Read();
                //            rowReader.Read();
                //            rowReader.Read();
                //        },
                //    }
                //});
                var dataset = reader.AsDataSet();
                foreach (DataTable table in dataset.Tables)
                {
                    var name = table.TableName;

                    int colCount = table.Columns.Count;
                    int rowCount = table.Rows.Count;
                    int count = 0;

                    //RemoveEmptyRows(table);
                    RemoveEmptyColumn(table);
                }
            }

            return Ok($"Excel/{newFileName}");
        }
        private static void RemoveEmptyRows(DataTable usedRange)
        {
            int count;
            DataTable curRange = new DataTable();

            count = usedRange.Rows.Count;

            for (int i = count; i > 0; i--)
            {
                bool isEmpty = true;

                var currenntcolumns = usedRange.Rows[i];
                var curRangeColumns = curRange.Rows[i];

                curRangeColumns = currenntcolumns;

                foreach (DataRow cell in curRangeColumns.Table.Rows)
                {
                    if (cell != null)
                    {
                        isEmpty = false;
                        break; // we can exit this loop since the range is not empty
                    }
                    else
                    {
                        // Cell value is null contiue checking
                    }
                } // end loop thru each cell in this range (row or column)

                if (isEmpty)
                {
                    curRangeColumns.Table.Rows.RemoveAt(i);
                }
            }
        }
        private static void RemoveEmptyColumn(DataTable usedRange)
        {
            int count;
            DataTable curRange = new DataTable();

            count = usedRange.Columns.Count;

            for (int i = count; i > 0; i--)
            {
                bool isEmpty = true;

                var currenntcolumns = usedRange.Columns[i];
                var curRangeColumns = curRange.Columns[i];

                curRangeColumns = currenntcolumns;

                foreach (DataColumn cell in curRangeColumns.Table.Columns)
                {
                    if (cell != null)
                    {
                        isEmpty = false;
                        break; // we can exit this loop since the range is not empty
                    }
                    else
                    {
                        // Cell value is null contiue checking
                    }
                } // end loop thru each cell in this range (row or column)

                if (isEmpty)
                {
                    curRangeColumns.Table.Columns.RemoveAt(i);
                }
            }
        }
    }
}
