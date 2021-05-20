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

            ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();

            int colCount = worksheet.Dimension.End.Column;  //get Column Count
            int rowCount = worksheet.Dimension.End.Row;     //get row count
            var objectDataValuesExcel = worksheet.Cells.Value;
            var contContainText = 0;

            for (int row = 1; row <= rowCount; row++)
            {
                //if (worksheet.Cells[row, 2].Value == null)
                //{
                //    worksheet.DeleteRow(row);
                //}
                for (int col = 1; col <= colCount; col++)
                {
                    var FirstCellValue = worksheet.Cells[row, col].Value?.ToString();
                    if (FirstCellValue != null)
                        contContainText += 1;

                    if (contContainText == 0)
                    {
                        worksheet.DeleteRow(row);
                        worksheet.DeleteColumn(col);
                        var stateWorksheet = worksheet.Cells.Value;
                        contContainText = 0;
                        break;
                    }
                    var stateWorksheet2 = worksheet.Cells.Value;
                }
            }

            return Ok($"Excel/{newFileName}");
        }
    }
}
