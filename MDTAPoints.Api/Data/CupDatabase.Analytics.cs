using MDTAPoints.Core;
using Microsoft.Data.SqlClient;

namespace MDTAPoints.Api.Data;

public sealed partial class CupDatabase
{
    public async Task<IReadOnlyList<ParticipationEntry>> LoadParticipationAsync(bool includeLowerDivisions, CancellationToken cancellationToken)
    {
        const string sql = """
            SELECT e.EventKind, en.Id, en.SchoolId, s.Name, c.StudentId, c.NormalizedName,
                   CAST(CASE WHEN studentSchool.ExcludeFromStats = 1 THEN 1 ELSE 0 END AS bit)
            FROM dbo.Tournaments t
            INNER JOIN dbo.Events e ON e.TournamentId = t.Id
            INNER JOIN dbo.Entries en ON en.EventId = e.Id
            INNER JOIN dbo.Schools s ON s.Id = en.SchoolId
            LEFT JOIN dbo.EntryCompetitors c ON c.EntryId = en.Id
            LEFT JOIN dbo.Students st ON st.Id = c.StudentId
            LEFT JOIN dbo.Schools studentSchool ON studentSchool.Id = st.SchoolId
            WHERE t.Status = N'Published'
              AND (@IncludeLower = 1 OR e.CountedDivision = 1)
              AND s.ExcludeFromStats = 0
            ORDER BY en.Id, c.Id
            """;

        await using var connection = await OpenAsync(cancellationToken);
        await using var command = new SqlCommand(sql, connection);
        command.Parameters.AddWithValue("@IncludeLower", includeLowerDivisions);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        var grouped = new Dictionary<int, MutableParticipation>();
        while (await reader.ReadAsync(cancellationToken))
        {
            if (!Enum.TryParse<EventKind>(reader.GetString(0), out var kind))
            {
                continue;
            }

            var entryId = reader.GetInt32(1);
            if (!grouped.TryGetValue(entryId, out var entry))
            {
                entry = new MutableParticipation(kind, entryId, reader.GetInt32(2), reader.GetString(3));
                grouped.Add(entryId, entry);
            }

            if (reader.IsDBNull(4) && reader.IsDBNull(5))
            {
                continue;
            }

            entry.SawCompetitor = true;
            if (reader.GetBoolean(6))
            {
                continue;
            }

            entry.Competitors.Add(new ParticipationCompetitor(
                reader.IsDBNull(4) ? null : reader.GetInt32(4),
                reader.IsDBNull(5) ? "" : reader.GetString(5)));
        }

        return grouped.Values
            .Where(entry => !entry.SawCompetitor || entry.Competitors.Count > 0)
            .Select(entry => entry.ToRecord())
            .ToList();
    }

    public async Task<IReadOnlyList<WeekendFieldEntry>> LoadWeekendFieldsAsync(CancellationToken cancellationToken)
    {
        const string sql = """
            SELECT t.Id, t.Name, t.StartDate, e.Name, e.EventKind, e.Division, COUNT(*)
            FROM dbo.Tournaments t
            INNER JOIN dbo.Events e ON e.TournamentId = t.Id
            INNER JOIN dbo.Entries en ON en.EventId = e.Id
            INNER JOIN dbo.Schools s ON s.Id = en.SchoolId
            WHERE t.Status = N'Published'
              AND s.ExcludeFromStats = 0
            GROUP BY t.Id, t.Name, t.StartDate, e.Id, e.Name, e.EventKind, e.Division
            """;

        await using var connection = await OpenAsync(cancellationToken);
        await using var command = new SqlCommand(sql, connection);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        var fields = new List<WeekendFieldEntry>();
        while (await reader.ReadAsync(cancellationToken))
        {
            if (!Enum.TryParse<EventKind>(reader.GetString(4), out var kind))
            {
                continue;
            }

            var eventName = reader.GetString(3);
            var divisionText = reader.IsDBNull(5) ? "" : reader.GetString(5);
            var division = Enum.TryParse<Division>(divisionText, true, out var parsed) && Enum.IsDefined(parsed)
                ? parsed
                : EventClassifier.ClassifyDivision(eventName);
            var count = reader.GetInt32(6);
            if (count <= 0)
            {
                continue;
            }

            fields.Add(new WeekendFieldEntry(
                reader.GetInt32(0),
                reader.GetString(1),
                DateOnly.FromDateTime(reader.GetDateTime(2)),
                kind,
                division,
                count));
        }

        return fields;
    }

    private sealed class MutableParticipation
    {
        public MutableParticipation(EventKind kind, int entryId, int schoolId, string schoolName)
        {
            Event = kind;
            EntryId = entryId;
            SchoolId = schoolId;
            SchoolName = schoolName;
        }

        public EventKind Event { get; }
        public int EntryId { get; }
        public int SchoolId { get; }
        public string SchoolName { get; }
        public bool SawCompetitor { get; set; }
        public List<ParticipationCompetitor> Competitors { get; } = new();

        public ParticipationEntry ToRecord() => new(Event, EntryId, SchoolId, SchoolName, Competitors);
    }
}
