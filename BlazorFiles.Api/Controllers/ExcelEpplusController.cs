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
    public class ExcelEpplusController : ControllerBase
    {

        private readonly IHostEnvironment _environment;

        public ExcelEpplusController(IHostEnvironment environment)
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


            return Ok($"Excel/{newFileName}");
        }
    }
}
