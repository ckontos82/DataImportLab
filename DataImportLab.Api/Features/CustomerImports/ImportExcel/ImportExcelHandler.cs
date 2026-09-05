using DataImportLab.Api.Features.CustomerImports.ImportExcel.Models;
using DataImportLab.Api.Infrastructure.Persistence;
using Microsoft.Data.SqlClient;

namespace DataImportLab.Api.Features.CustomerImports.ImportExcel;

public class ImportExcelHandler(
    CustomerExcelReader excelReader,
    ImportExcelJobRepository jobRepository,
    ImportExcelStagingRepository stagingRepository,
    SqlConnectionFactory connectionFactory,
    ILogger<ImportExcelHandler> logger)
{
    public async Task<ImportExcelResult> HandleAsync(
        Stream stream,
        string fileName,
        CancellationToken cancellationToken)
    {
        var jobId = await jobRepository.CreateAsync(fileName, cancellationToken);

        try
        {
            await using var connection = connectionFactory.CreateConnection();
            await connection.OpenAsync(cancellationToken);
            await using var transaction =
                (SqlTransaction)await connection.BeginTransactionAsync(cancellationToken);

            var totalRows = await stagingRepository.WriteAsync(
                connection, transaction, jobId,
                excelReader.ReadRows(stream), cancellationToken);

            await jobRepository.MarkStagedAsync(
                connection, transaction, jobId, totalRows, cancellationToken);

            await transaction.CommitAsync(cancellationToken);

            return new ImportExcelResult { JobId = jobId, TotalRows = totalRows };
        }
        catch (Exception exception)
        {
            // Disposing the uncommitted transaction above rolls back every batch.
            logger.LogError(exception, "Excel import job {JobId} failed.", jobId);

            try
            {
                // Record failure even if the HTTP request was cancelled.
                using var failureTimeout = new CancellationTokenSource(TimeSpan.FromSeconds(10));
                await jobRepository.MarkFailedAsync(jobId, failureTimeout.Token);
            }
            catch (Exception statusException)
            {
                logger.LogError(statusException,
                    "Could not mark Excel import job {JobId} as failed.", jobId);
            }

            throw;
        }
    }
}
