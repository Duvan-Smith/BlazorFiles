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
        //private static void ConventionalRemoveEmptyRowsCols(ExcelWorksheet worksheet)
        //{
        //    ExcelRange usedRange = worksheet.Cells;

        //    int colCount = worksheet.Dimension.End.Column;  //get Column Count
        //    int rowCount = worksheet.Dimension.End.Row;     //get row count

        //    //int totalRows = usedRange.Rows.Count;
        //    //int totalCols = usedRange.Columns.Count;

        //    RemoveEmpty(colCount, rowCount, RowOrCol.Row);
        //    RemoveEmpty(colCount, rowCount, RowOrCol.Column);
        //}

        //private static void RemoveEmpty(int colCount, int rowCount, RowOrCol rowOrCol)
        //{
        //    int count;
        //    //ExcelRange curRange;
        //    List<ExcelRange> curRangeList = new List<ExcelRange>();
        //    //if (rowOrCol == RowOrCol.Column)
        //    //    count = usedRange.Columns;
        //    //else
        //    //    count = usedRange.Rows;

        //    for (int row = 1; row <= rowCount; row++)
        //    {
        //        for (int col = 1; col <= colCount; col++)
        //        {
        //            Console.WriteLine(" Row:" + row + " column:" + col + " Value:" + worksheet.Cells[row, col].Value?.ToString().Trim());
        //        }
        //    }

        //    //for (int i = count; i > 0; i--)
        //    //{
        //    //    bool isEmpty = true;
        //    //    if (rowOrCol == RowOrCol.Column)
        //    //        curRange = usedRange.Columns[i];
        //    //    else
        //    //        curRange = usedRange.Rows[i];

        //    //    foreach (Excel.Range cell in curRange.Cells)
        //    //    {
        //    //        if (cell.Value != null)
        //    //        {
        //    //            isEmpty = false;
        //    //            break; // we can exit this loop since the range is not empty
        //    //        }
        //    //        else
        //    //        {
        //    //            // Cell value is null contiue checking
        //    //        }
        //    //    } // end loop thru each cell in this range (row or column)

        //        if (isEmpty)
        //        {
        //            curRange.Delete(eShiftTypeDelete.Left);
        //        }
        //    }
        //}

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
                for (int col = 1; col <= colCount; col++)
                {
                    var FirstCellValue = worksheet.Cells[row, col].Value?.ToString().Trim();
                    if (worksheet.Cells[row, col].Value == null)
                    {
                        //TODO: Lenny y Michael, validar que se eliminen filas y columnas
                        //vacias, remplazar el metodo de Delerow que esta abajo
                        //solution : cambiar el contador de colcount a uno para que 
                        //solo lo haga una vez
                        //worksheet.DeleteRow(row, col, true);
                        var stateWorksheet = worksheet.Cells.Value;
                    }
                }
            }

            return Ok($"Excel/{newFileName}");
        }
    }
}
