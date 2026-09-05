namespace DataImportLab.Api.Features.CustomerImports.ImportExcel.Models;

public sealed record ImportExcelResult
{
    public long JobId { get; init; }
    public int TotalRows { get; init; }
    public string Status { get; init; } = "PENDING";
}
