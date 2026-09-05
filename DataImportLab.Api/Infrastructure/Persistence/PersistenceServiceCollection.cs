namespace DataImportLab.Api.Infrastructure.Persistence;

public static class PersistenceServiceCollection
{
    public static IServiceCollection AddPersistence(
        this IServiceCollection services,
        IConfiguration configuration)
    {
        var connectionString =
            configuration.GetConnectionString("ImportDb");

        if (string.IsNullOrWhiteSpace(connectionString))
        {
            throw new InvalidOperationException(
                "Λείπει το connection string ImportDb.");
        }

        services.AddSingleton(
            new SqlConnectionFactory(connectionString));

        return services;
    }
}
