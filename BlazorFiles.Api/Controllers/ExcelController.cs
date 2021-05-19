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

                Console.WriteLine("Entrada 1");
                var reader = ExcelReaderFactory.CreateReader(fileStream,
                new ExcelReaderConfiguration() { FallbackEncoding = System.Text.Encoding.GetEncoding(1252) });
                var dataset = reader.AsDataSet(new ExcelDataSetConfiguration()
                {
                    ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                    {
                        UseHeaderRow = true
                    }
                });
                foreach (DataTable table in dataset.Tables)
                {
                    foreach (DataRow row in table.Rows)
                    {
                        var apellido = (string)row["Apellido"];
                        Console.WriteLine((string)row["Apellido"]);
                        Console.WriteLine((string)row["Nombre"]);
                        Console.WriteLine((string)row["Aficion"]);
                        //personasLista.Add(new Persona()
                        //{
                        //    Apellido = (string)row["Apellido"],
                        //    Nombre = (string)row["Nombre"],
                        //    Aficion = (string)row["Aficion"]
                        //});
                    }
                }
            }
            return Ok($"Excel/{newFileName}");
        }
    }
}
