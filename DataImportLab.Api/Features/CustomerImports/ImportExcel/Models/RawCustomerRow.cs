namespace DataImportLab.Api.Features.CustomerImports.ImportExcel.Models;

public sealed record RawCustomerRow
{
    public int SourceRow { get; init; }
    public string? CustomerId { get; init; }
    public string? FirstName { get; init; }
    public string? LastName { get; init; }
    public string? Email { get; init; }
    public string? VatNumber { get; init; }
    public string? Status { get; init; }
    public string? Balance { get; init; }
    public string? CreatedAt { get; init; }
    public string? UpdatedAt { get; init; }
}