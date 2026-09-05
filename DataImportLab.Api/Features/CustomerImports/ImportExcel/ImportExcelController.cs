using Microsoft.AspNetCore.Mvc;

namespace DataImportLab.Api.Features.CustomerImports.ImportExcel;

[ApiController]
[Route("api/imports/customers")]
public class ImportExcelController(CustomerExcelReader excelReader) : ControllerBase
{
    [HttpPost("excel")]
    [Consumes("multipart/form-data")]
    public IActionResult Import(IFormFile file)
    {
        if (file.Length == 0)
            return BadRequest("Το αρχείο είναι κενό.");

        try
        {
            using var stream = file.OpenReadStream();
            var totalRows = excelReader.ReadRows(stream).Count();

            return Ok(new
            {
                fileName = file.FileName,
                sizeInBytes = file.Length,
                totalRows
            });
        }
        catch (InvalidExcelImportException ex)
        {
            return Problem(
                statusCode: StatusCodes.Status400BadRequest,
                title: "Μη έγκυρο αρχείο Excel",
                detail: ex.Message);
        }
    }
}

