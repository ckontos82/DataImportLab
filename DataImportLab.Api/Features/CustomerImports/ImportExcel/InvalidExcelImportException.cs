namespace DataImportLab.Api.Features.CustomerImports.ImportExcel;

public sealed class InvalidExcelImportException : Exception
{
    public InvalidExcelImportException(string message)
        : base(message)
    {
    }
}
