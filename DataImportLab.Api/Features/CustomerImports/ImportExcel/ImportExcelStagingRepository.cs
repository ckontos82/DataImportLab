using System.Data;
using DataImportLab.Api.Features.CustomerImports.ImportExcel.Models;
using Microsoft.Data.SqlClient;

namespace DataImportLab.Api.Features.CustomerImports.ImportExcel;

public sealed class ImportExcelStagingRepository
{
    private const int BatchSize = 1000;

    public async Task<int> WriteAsync(
        SqlConnection connection,
        SqlTransaction transaction,
        long jobId,
        IEnumerable<RawCustomerRow> rows,
        CancellationToken cancellationToken)
    {
        using var batch = CreateBatchTable();
        using var bulkCopy = new SqlBulkCopy(
            connection, SqlBulkCopyOptions.CheckConstraints, transaction)
        {
            DestinationTableName = "dbo.staging_customers",
            BatchSize = BatchSize
        };

        foreach (DataColumn column in batch.Columns)
            bulkCopy.ColumnMappings.Add(column.ColumnName, column.ColumnName);

        var totalRows = 0;

        foreach (var row in rows)
        {
            cancellationToken.ThrowIfCancellationRequested();

            batch.Rows.Add(
                jobId,
                row.SourceRow,
                ToStagingValue(row.CustomerId, "customer_id", row.SourceRow, 100),
                ToStagingValue(row.FirstName, "first_name", row.SourceRow),
                ToStagingValue(row.LastName, "last_name", row.SourceRow),
                ToStagingValue(row.Email, "email", row.SourceRow),
                ToStagingValue(row.VatNumber, "vat_number", row.SourceRow),
                ToStagingValue(row.Status, "status", row.SourceRow),
                ToStagingValue(row.Balance, "balance", row.SourceRow),
                ToStagingValue(row.CreatedAt, "created_at", row.SourceRow),
                ToStagingValue(row.UpdatedAt, "updated_at", row.SourceRow));

            totalRows++;

            if (batch.Rows.Count == BatchSize)
            {
                await bulkCopy.WriteToServerAsync(batch, cancellationToken);
                batch.Clear();
            }
        }

        if (batch.Rows.Count > 0)
            await bulkCopy.WriteToServerAsync(batch, cancellationToken);

        return totalRows;
    }

    private static DataTable CreateBatchTable()
    {
        var table = new DataTable();
        table.Columns.Add("job_id", typeof(long));
        table.Columns.Add("source_row", typeof(int));
        table.Columns.Add("customer_id", typeof(string));
        table.Columns.Add("first_name", typeof(string));
        table.Columns.Add("last_name", typeof(string));
        table.Columns.Add("email", typeof(string));
        table.Columns.Add("vat_number", typeof(string));
        table.Columns.Add("status", typeof(string));
        table.Columns.Add("balance", typeof(string));
        table.Columns.Add("created_at", typeof(string));
        table.Columns.Add("updated_at", typeof(string));
        return table;
    }

    private static object ToStagingValue(
        string? value, string column, int sourceRow, int maxLength = 4000)
    {
        // Reject values that cannot fit in staging instead of truncating raw data.
        if (value?.Length > maxLength)
        {
            throw new InvalidExcelImportException(
                $"Στη γραμμή {sourceRow}, το πεδίο {column} υπερβαίνει " +
                $"το όριο των {maxLength} χαρακτήρων του staging.");
        }

        return (object?)value ?? DBNull.Value;
    }
}
