using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Hosting;
using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace BlazorFiles.Api.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExcelInteropController : ControllerBase
    {
        private readonly IHostEnvironment _environment;

        public ExcelInteropController(IHostEnvironment environment)
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

            //var app = new Microsoft.Office.Interop.Excel.Application();
            //var book = app.Workbooks.Open(Filename: filePath, Format: Microsoft.Office.Interop.Excel.XlFileFormat.xlCSV);

            //foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in book.Sheets)
            //{
            //    DataTable dt = new DataTable(sheet.Name);
            //    //??? Fill dt from sheet 
            //}

            //Microsoft.Office.Interop.Excel.Application Excel = new Microsoft.Office.Interop.Excel.Application();
            //string originalPath = newFileName;
            //Microsoft.Office.Interop.Excel.Workbook workbook = Excel.Workbooks.Open(originalPath);
            //Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets["colores"];
            //Microsoft.Office.Interop.Excel.Range usedRange = worksheet.UsedRange;

            return Ok($"Excel/{newFileName}");
        }
    }
}
