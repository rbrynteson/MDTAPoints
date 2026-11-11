using Microsoft.Data.SqlClient;

namespace MDTAPoints.Api.Data;

public sealed partial class CupDatabase
{
    public async Task<IReadOnlyList<SchoolListItem>> ListSchoolsAsync(CancellationToken cancellationToken)
    {
        const string sql = """
            SELECT s.Id, s.Name, s.ExcludeFromStats,
                   (SELECT COUNT(*) FROM dbo.Entries en WHERE en.SchoolId = s.Id)
            FROM dbo.Schools s
            ORDER BY s.Name
            """;
        await using var connection = await OpenAsync(cancellationToken);
        await using var command = new SqlCommand(sql, connection);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        var schools = new List<SchoolListItem>();
        while (await reader.ReadAsync(cancellationToken))
        {
            schools.Add(ReadSchool(reader));
        }

        return schools;
    }

    public async Task<SchoolListItem?> GetSchoolAsync(int id, CancellationToken cancellationToken)
    {
        const string sql = """
            SELECT s.Id, s.Name, s.ExcludeFromStats,
                   (SELECT COUNT(*) FROM dbo.Entries en WHERE en.SchoolId = s.Id)
            FROM dbo.Schools s
            WHERE s.Id = @Id
            """;
        await using var connection = await OpenAsync(cancellationToken);
        await using var command = new SqlCommand(sql, connection);
        command.Parameters.AddWithValue("@Id", id);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        return await reader.ReadAsync(cancellationToken) ? ReadSchool(reader) : null;
    }

    public async Task<bool> SetSchoolExcludedAsync(int id, bool excluded, CancellationToken cancellationToken)
    {
        await using var connection = await OpenAsync(cancellationToken);
        await using var command = new SqlCommand(
            "UPDATE dbo.Schools SET ExcludeFromStats = @Excluded WHERE Id = @Id",
            connection);
        command.Parameters.AddWithValue("@Excluded", excluded);
        command.Parameters.AddWithValue("@Id", id);
        return await command.ExecuteNonQueryAsync(cancellationToken) == 1;
    }

    private static SchoolListItem ReadSchool(SqlDataReader reader) => new()
    {
        Id = reader.GetInt32(0),
        Name = reader.GetString(1),
        ExcludeFromStats = reader.GetBoolean(2),
        EntryCount = reader.GetInt32(3)
    };
}
