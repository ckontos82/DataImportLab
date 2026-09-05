using DataImportLab.Api.Features.CustomerImports.ImportExcel.Models;
using ExcelDataReader;
using System.Globalization;

namespace DataImportLab.Api.Features.CustomerImports.ImportExcel;

public class CustomerExcelReader
{
    private static readonly string[] ExpectedHeaders =
    [
        "customer_id",
        "first_name",
        "last_name",
        "email",
        "vat_number",
        "status",
        "balance",
        "created_at",
        "updated_at"
    ];

    public IEnumerable<RawCustomerRow> ReadRows(Stream stream)
    {
        using var reader = ExcelReaderFactory.CreateOpenXmlReader(
            stream, new ExcelReaderConfiguration { LeaveOpen = true });

        while (!string.Equals(reader.Name, "Customers", StringComparison.OrdinalIgnoreCase))
        {
            if (!reader.NextResult())
                throw new InvalidExcelImportException(
                    "Δεν βρέθηκε το φύλλο Customers.");
        }

        // Skip the header only after locating the Customers sheet.
        if (!reader.Read())
            throw new InvalidExcelImportException(
                "Το φύλλο Customers είναι κενό.");
        
        ValidateHeaders(reader);

        var sourceRow = 1;

        while (reader.Read())
        {
            sourceRow++;

            yield return new RawCustomerRow
            {
                SourceRow = sourceRow,
                CustomerId = ToRawText(reader.GetValue(0)),
                FirstName = ToRawText(reader.GetValue(1)),
                LastName = ToRawText(reader.GetValue(2)),
                Email = ToRawText(reader.GetValue(3)),
                VatNumber = ToRawText(reader.GetValue(4)),
                Status = ToRawText(reader.GetValue(5)),
                Balance = ToRawText(reader.GetValue(6)),
                CreatedAt = ToRawText(reader.GetValue(7)),
                UpdatedAt = ToRawText(reader.GetValue(8))
            };
        }
    }

    private static string? ToRawText(object? value)
    {
        return value switch
        {
            null => null,
            DateTime date => date.ToString(
                "yyyy-MM-dd", CultureInfo.InvariantCulture),
            IFormattable formatted => formatted.ToString(
                null, CultureInfo.InvariantCulture),
            _ => value.ToString()
        };
    }

    private static void ValidateHeaders(IExcelDataReader reader)
    {
        if (reader.FieldCount != ExpectedHeaders.Length)
        {
            throw new InvalidExcelImportException(
                $"Αναμένονται {ExpectedHeaders.Length} στήλες, " +
                $"αλλά βρέθηκαν {reader.FieldCount}.");
        }

        for (var i = 0; i < ExpectedHeaders.Length; i++)
        {
            var actual = reader.GetValue(i)?.ToString()?.Trim();

            if (!string.Equals(
                actual, ExpectedHeaders[i],
                StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidExcelImportException(
                    $"Στη στήλη {i + 1} αναμένεται το header {ExpectedHeaders[i]}, αλλά βρέθηκε {actual}.");
            }
        }
    }
}
