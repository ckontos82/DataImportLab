using System.Data;
using DataImportLab.Api.Infrastructure.Persistence;
using Microsoft.Data.SqlClient;

namespace DataImportLab.Api.Features.CustomerImports.ImportExcel;

public sealed class ImportExcelJobRepository(SqlConnectionFactory connectionFactory)
{
    public async Task<long> CreateAsync(
        string sourceFile,
        CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(sourceFile) || sourceFile.Length > 255)
        {
            throw new InvalidExcelImportException(
                "Το όνομα αρχείου πρέπει να έχει από 1 έως 255 χαρακτήρες.");
        }

        const string sql = """
            INSERT INTO dbo.import_jobs (source_file, source_type, status, started_at)
            OUTPUT INSERTED.job_id
            VALUES (@sourceFile, 'XLSX', 'RUNNING', SYSUTCDATETIME());
            """;

        await using var connection = connectionFactory.CreateConnection();
        await connection.OpenAsync(cancellationToken);

        using var command = connection.CreateCommand();
        command.CommandText = sql;
        command.Parameters.Add("@sourceFile", SqlDbType.NVarChar, 255).Value = sourceFile;

        var result = await command.ExecuteScalarAsync(cancellationToken);

        return result is long jobId
            ? jobId
            : throw new InvalidOperationException(
                "Η δημιουργία του import job δεν επέστρεψε job_id.");
    }

    public async Task MarkStagedAsync(
        SqlConnection connection,
        SqlTransaction transaction,
        long jobId,
        int totalRows,
        CancellationToken cancellationToken)
    {
        // Customer validation is a later step, so the import remains pending.
        const string sql = """
            UPDATE dbo.import_jobs
            SET status = 'PENDING', total_rows = @totalRows
            WHERE job_id = @jobId;
            """;

        using var command = connection.CreateCommand();
        command.Transaction = transaction;
        command.CommandText = sql;
        command.Parameters.Add("@jobId", SqlDbType.BigInt).Value = jobId;
        command.Parameters.Add("@totalRows", SqlDbType.Int).Value = totalRows;

        if (await command.ExecuteNonQueryAsync(cancellationToken) != 1)
            throw new InvalidOperationException("Δεν βρέθηκε το import job για ενημέρωση.");
    }

    public async Task MarkFailedAsync(long jobId, CancellationToken cancellationToken)
    {
        const string sql = """
            UPDATE dbo.import_jobs
            SET status = 'FAILED', completed_at = SYSUTCDATETIME()
            WHERE job_id = @jobId;
            """;

        await using var connection = connectionFactory.CreateConnection();
        await connection.OpenAsync(cancellationToken);

        using var command = connection.CreateCommand();
        command.CommandText = sql;
        command.Parameters.Add("@jobId", SqlDbType.BigInt).Value = jobId;

        if (await command.ExecuteNonQueryAsync(cancellationToken) != 1)
            throw new InvalidOperationException("Δεν βρέθηκε το import job για ενημέρωση.");
    }
}
