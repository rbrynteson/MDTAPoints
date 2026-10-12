using Microsoft.Data.SqlClient;

namespace MDTAPoints.Api.Data;

public sealed class SchemaMigrator
{
    private readonly string _connectionString;

    public SchemaMigrator(IConfiguration configuration)
    {
        _connectionString = configuration.GetConnectionString("Cup")
            ?? throw new InvalidOperationException("Set ConnectionStrings:Cup to the Azure SQL database.");
        if (string.IsNullOrWhiteSpace(_connectionString))
        {
            throw new InvalidOperationException("Set ConnectionStrings:Cup to the Azure SQL database.");
        }
    }

    public async Task ApplyAsync(CancellationToken cancellationToken)
    {
        await using var connection = new SqlConnection(_connectionString);
        await connection.OpenAsync(cancellationToken);
        var applied = await LoadAppliedAsync(connection, cancellationToken);

        foreach (var script in ReadScripts())
        {
            if (applied.Contains(script.Version))
            {
                continue;
            }

            await using (var command = new SqlCommand(script.Sql, connection))
            {
                await command.ExecuteNonQueryAsync(cancellationToken);
            }

            await using var mark = new SqlCommand(
                "INSERT INTO dbo.SchemaVersion (Version, AppliedAt) VALUES (@Version, SYSUTCDATETIME())",
                connection);
            mark.Parameters.AddWithValue("@Version", script.Version);
            await mark.ExecuteNonQueryAsync(cancellationToken);
        }
    }

    private static async Task<HashSet<int>> LoadAppliedAsync(SqlConnection connection, CancellationToken cancellationToken)
    {
        await using var exists = new SqlCommand("SELECT OBJECT_ID(N'dbo.SchemaVersion', N'U')", connection);
        var id = await exists.ExecuteScalarAsync(cancellationToken);
        if (id is null or DBNull)
        {
            return new HashSet<int>();
        }

        var versions = new HashSet<int>();
        await using var command = new SqlCommand("SELECT Version FROM dbo.SchemaVersion", connection);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        while (await reader.ReadAsync(cancellationToken))
        {
            versions.Add(reader.GetInt32(0));
        }

        return versions;
    }

    private static IEnumerable<(int Version, string Sql)> ReadScripts()
    {
        var assembly = typeof(SchemaMigrator).Assembly;
        foreach (var name in assembly.GetManifestResourceNames().Where(resource => resource.EndsWith(".sql", StringComparison.OrdinalIgnoreCase)).OrderBy(resource => resource))
        {
            using var stream = assembly.GetManifestResourceStream(name) ?? throw new InvalidOperationException($"Missing SQL script {name}.");
            using var reader = new StreamReader(stream);
            var file = name.Split('.')[^2];
            var version = int.Parse(file.Split('_')[0]);
            yield return (version, reader.ReadToEnd());
        }
    }
}
