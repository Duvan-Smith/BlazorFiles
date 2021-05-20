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

            for (int row = 1; row <= rowCount; row++)
            {

                if (worksheet.Cells[row, 2].Value == null)
                {
                    worksheet.DeleteRow(row);
                }
                for (int col = 1; col <= colCount; col++)
                {

                    var FirstCellValue = worksheet.Cells[row, col].Value?.ToString().Trim();
                    if (FirstCellValue == null)
                    {
                        //TODO: Lenny y Michael, validar que se eliminen las columnas
                        //vacias, remplazar el metodo de DeleteColumn que esta abajo
                        //solution : Colocar un condicional para que solo elimine una sola columna a la vez.
                        //worksheet.DeleteRow(row, col, true);
                        //worksheet.DeleteColumn(col);
                        var stateWorksheet = worksheet.Cells.Value;
                    }
                }
            }

            return Ok($"Excel/{newFileName}");
        }
    }
}
