using Microsoft.Data.SqlClient;

namespace DataImportLab.Api.Infrastructure.Persistence;

public sealed class SqlConnectionFactory(string connectionString)
{
    public SqlConnection CreateConnection()
    {
        return new SqlConnection(connectionString);
    }
}
