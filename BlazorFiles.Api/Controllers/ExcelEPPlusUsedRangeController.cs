using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Hosting;
using OfficeOpenXml;
using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace BlazorFiles.Api.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExcelEPPlusUsedRangeController : ControllerBase
    {
        //
        enum RowOrCol { Row, Column };

        private readonly IHostEnvironment _environment;

        public ExcelEPPlusUsedRangeController(IHostEnvironment environment)
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

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var package = new ExcelPackage(new FileInfo(filePath));

            foreach (ExcelWorksheet worksheet in package.Workbook.Worksheets)
            {
                int colCount = worksheet.Dimension.End.Column;
                int rowCount = worksheet.Dimension.End.Row;
                var objectDataValuesExcel = worksheet.Cells.Value;
                int count = 0;
                for (int row = 1; row <= rowCount; row++)
                {
                    for (int col = 1; col <= colCount; col++)
                    {
                        if (worksheet.Cells[row, col].Value?.ToString().Trim() == null)
                        {
                            count++;
                        }
                        if (count == colCount)
                        {
                            worksheet.DeleteRow(row);
                        }
                        objectDataValuesExcel = worksheet.Cells.Value;
                    }
                    count = 0;
                }
                for (int col = 1; col <= colCount; col++)
                {
                    for (int row = 1; row <= rowCount; row++)
                    {
                        if (worksheet.Cells[row, col].Value?.ToString().Trim() == null)
                        {
                            count++;
                        }
                        if (count == rowCount)
                        {
                            worksheet.DeleteColumn(col);
                        }
                        objectDataValuesExcel = worksheet.Cells.Value;
                    }
                    count = 0;
                }
                Console.WriteLine(objectDataValuesExcel);
            }

            return Ok($"Excel/{newFileName}");
        }
    }
}
