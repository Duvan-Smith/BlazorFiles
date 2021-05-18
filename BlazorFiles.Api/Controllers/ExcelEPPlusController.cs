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
    public class ExcelEPPlusController : ControllerBase
    {

        private readonly IHostEnvironment _environment;

        public ExcelEPPlusController(IHostEnvironment environment)
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

            var package = new ExcelPackage(new FileInfo(@filePath));
            ExcelWorksheet sheet = package.Workbook.Worksheets[1];

            var table = sheet.Tables.First();

            //FileInfo existingFile = new FileInfo(filePath);
            //using (ExcelPackage package = new ExcelPackage(existingFile))
            //{
            //    //get the first worksheet in the workbook
            //    ExcelWorksheet worksheet = package.Workbook.Worksheets[1];
            //    int colCount = worksheet.Dimension.End.Column;  //get Column Count
            //    int rowCount = worksheet.Dimension.End.Row;     //get row count
            //    for (int row = 1; row <= rowCount; row++)
            //    {
            //        for (int col = 1; col <= colCount; col++)
            //        {
            //            Console.WriteLine(" Row:" + row + " column:" + col + " Value:" + worksheet.Cells[row, col].Value?.ToString().Trim());
            //        }
            //    }
            //}

            //FileInfo fileInfo = new FileInfo(filePath);

            //ExcelPackage package = new ExcelPackage(fileInfo);

            //ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();

            //int rows = worksheet.Dimension.Rows; // 20
            //int columns = worksheet.Dimension.Columns; // 7

            //// loop through the worksheet rows and columns
            //for (int i = 1; i <= rows; i++)
            //{
            //    for (int j = 1; j <= columns; j++)
            //    {

            //        string content = worksheet.Cells[i, j].Value.ToString();
            //        /* Do something ...*/
            //    }
            //}

            return Ok($"Excel/{newFileName}");
        }
    }
}
