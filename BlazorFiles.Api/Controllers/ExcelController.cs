using BlazorFiles.Api.TablasParametricasDto;
using ExcelDataReader;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Hosting;
using System;
using System.Collections.Generic;
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
                List<string> listOne = new List<string>();

                List<DataTransferObject> ListDataT = new List<DataTransferObject>();
                List<TablaParametricaDto> listTablaP = new List<TablaParametricaDto>();

                List<MarcasDto> listMarcas = new List<MarcasDto>();

                foreach (DataTable table in dataset.Tables)
                {
                    listOne.Add(table.TableName);
                    foreach (DataRow row in table.Rows)
                    {
                        var apellido = (string)row["Apellido"];
                        listTablaP.Add(new MarcasDto()
                        {
                            Title = (string)row["apellido"],
                            CodigoMarca = (string)row["nombre"],
                            DescripcionMarca = (string)row["aficion"]
                        });
                    }
                }
            }

            return Ok($"Excel/{newFileName}");
        }
    }
}
