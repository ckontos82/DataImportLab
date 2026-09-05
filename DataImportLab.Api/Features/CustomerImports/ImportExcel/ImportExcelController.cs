using Microsoft.AspNetCore.Mvc;

namespace DataImportLab.Api.Features.CustomerImports.ImportExcel;

[ApiController]
[Route("api/imports/customers")]
public class ImportExcelController(ImportExcelHandler handler) : ControllerBase
{
    [HttpPost("excel")]
    [Consumes("multipart/form-data")]
    public async Task<IActionResult> Import(IFormFile file, CancellationToken cancellationToken)
    {
        if (file.Length == 0)
            return BadRequest("Το αρχείο είναι κενό.");

        try
        {
            using var stream = file.OpenReadStream();
            var result = await handler.HandleAsync(stream, file.FileName, cancellationToken);

            return Ok(new
            {
                fileName = file.FileName,
                sizeInBytes = file.Length,
                jobId = result.JobId,
                totalRows = result.TotalRows,
                status = result.Status
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

