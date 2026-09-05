using DataImportLab.Api.Features.CustomerImports.ImportExcel;

namespace DataImportLab.Api.Features.CustomerImports;

public static class CustomerImportsServiceCollection
{
    public static IServiceCollection AddCustomerImports(this IServiceCollection services)
    {
        System.Text.Encoding.RegisterProvider(
            System.Text.CodePagesEncodingProvider.Instance);

        services.AddScoped<CustomerExcelReader>();

        return services;
    }
}
