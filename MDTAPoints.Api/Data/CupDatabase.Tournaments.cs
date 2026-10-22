using MDTAPoints.Core;
using MDTAPoints.Tabroom;
using Microsoft.Data.SqlClient;

namespace MDTAPoints.Api.Data;

public sealed class ReviewUpdate
{
    public string Kind { get; set; } = TournamentKind.Invitational.ToString();
    public Dictionary<int, bool> Octos { get; } = new();
    public Dictionary<int, bool> FinalSessions { get; } = new();
    public Dictionary<int, bool> Semis { get; } = new();
    public Dictionary<int, EntryEdit> Entries { get; } = new();
    public Dictionary<int, string> CompetitorChoices { get; } = new();
}

public sealed class EntryEdit
{
    public int ElimWins { get; set; }
    public bool ReachedElims { get; set; }
    public bool ReachedFinals { get; set; }
    public bool ReachedSemis { get; set; }
    public bool QualifiedToState { get; set; }
    public int? Place { get; set; }
    public string? PlaceNote { get; set; }
}

public sealed partial class CupDatabase
{
    public async Task<IReadOnlyList<TournamentListItem>> ListTournamentsAsync(CancellationToken cancellationToken)
    {
        const string sql = """
            SELECT Id, TabroomTournamentId, Name, StartDate, Kind, Status, InState
            FROM dbo.Tournaments
            ORDER BY StartDate DESC, Name
            """;
        await using var connection = await OpenAsync(cancellationToken);
        var items = new List<TournamentListItem>();
        await using var command = new SqlCommand(sql, connection);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        while (await reader.ReadAsync(cancellationToken))
        {
            items.Add(new TournamentListItem
            {
                Id = reader.GetInt32(0),
                TabroomTournamentId = reader.IsDBNull(1) ? null : reader.GetInt32(1),
                Name = reader.GetString(2),
                StartDate = DateOnly.FromDateTime(reader.GetDateTime(3)),
                Kind = reader.GetString(4),
                Status = reader.GetString(5),
                InState = reader.GetBoolean(6)
            });
        }

        return items;
    }

    public async Task<TournamentListItem?> FindByTabroomIdAsync(int tabroomId, CancellationToken cancellationToken)
    {
        var matches = await ListTournamentsAsync(cancellationToken);
        return matches.FirstOrDefault(item => item.TabroomTournamentId == tabroomId);
    }

    public async Task<TournamentDetail?> GetTournamentAsync(int id, CancellationToken cancellationToken)
    {
        await using var connection = await OpenAsync(cancellationToken);
        TournamentDetail? tournament = null;
        await using (var command = new SqlCommand("""
            SELECT Id, TabroomTournamentId, Name, StartDate, StateCode, InState, Kind, Status, SkippedSummary
            FROM dbo.Tournaments WHERE Id = @Id
            """, connection))
        {
            command.Parameters.AddWithValue("@Id", id);
            await using var reader = await command.ExecuteReaderAsync(cancellationToken);
            if (await reader.ReadAsync(cancellationToken))
            {
                tournament = new TournamentDetail
                {
                    Id = reader.GetInt32(0),
                    TabroomTournamentId = reader.IsDBNull(1) ? null : reader.GetInt32(1),
                    Name = reader.GetString(2),
                    StartDate = DateOnly.FromDateTime(reader.GetDateTime(3)),
                    StateCode = reader.IsDBNull(4) ? null : reader.GetString(4),
                    InState = reader.GetBoolean(5),
                    Kind = reader.GetString(6),
                    Status = reader.GetString(7),
                    SkippedSummary = reader.IsDBNull(8) ? null : reader.GetString(8)
                };
            }
        }

        if (tournament is null)
        {
            return null;
        }

        var events = new Dictionary<int, EventRow>();
        await using (var command = new SqlCommand("""
            SELECT Id, Name, EventKind, SchoolCount, EntryCount, BreaksToOctosOrDeeper, HasFinalSession, HasSemis, IsHouse, CountedDivision
            FROM dbo.Events WHERE TournamentId = @Id ORDER BY Name
            """, connection))
        {
            command.Parameters.AddWithValue("@Id", id);
            await using var reader = await command.ExecuteReaderAsync(cancellationToken);
            while (await reader.ReadAsync(cancellationToken))
            {
                var row = new EventRow
                {
                    Id = reader.GetInt32(0),
                    Name = reader.GetString(1),
                    EventKind = reader.GetString(2),
                    SchoolCount = reader.GetInt32(3),
                    EntryCount = reader.GetInt32(4),
                    BreaksToOctosOrDeeper = reader.GetBoolean(5),
                    HasFinalSession = reader.GetBoolean(6),
                    HasSemis = reader.GetBoolean(7),
                    IsHouse = reader.GetBoolean(8),
                    CountedDivision = reader.GetBoolean(9)
                };
                events.Add(row.Id, row);
            }
        }

        var entries = new Dictionary<int, EntryRow>();
        if (events.Count > 0)
        {
            await using var command = new SqlCommand("""
                SELECT en.Id, en.EventId, en.EntryName, en.SchoolId, s.Name, en.Place, en.IsPrelim,
                       en.ReachedElims, en.ElimWins, en.ReachedFinals, en.ReachedSemis, en.QualifiedToState,
                       en.TabroomPlace, en.PlaceNote
                FROM dbo.Entries en
                INNER JOIN dbo.Schools s ON s.Id = en.SchoolId
                INNER JOIN dbo.Events e ON e.Id = en.EventId
                WHERE e.TournamentId = @Id
                ORDER BY en.EventId, en.Place, en.EntryName
                """, connection);
            command.Parameters.AddWithValue("@Id", id);
            await using var reader = await command.ExecuteReaderAsync(cancellationToken);
            while (await reader.ReadAsync(cancellationToken))
            {
                var row = new EntryRow
                {
                    Id = reader.GetInt32(0),
                    EntryName = reader.GetString(2),
                    SchoolId = reader.GetInt32(3),
                    SchoolName = reader.GetString(4),
                    Place = reader.GetInt32(5),
                    IsPrelim = reader.GetBoolean(6),
                    ReachedElims = reader.GetBoolean(7),
                    ElimWins = reader.GetInt32(8),
                    ReachedFinals = reader.GetBoolean(9),
                    ReachedSemis = reader.GetBoolean(10),
                    QualifiedToState = reader.GetBoolean(11),
                    TabroomPlace = reader.IsDBNull(12) ? null : reader.GetInt32(12),
                    PlaceNote = reader.IsDBNull(13) ? null : reader.GetString(13)
                };
                entries.Add(row.Id, row);
                events[reader.GetInt32(1)].Entries.Add(row);
            }
        }

        if (entries.Count > 0)
        {
            await using var command = new SqlCommand("""
                SELECT c.Id, c.EntryId, c.RawName, c.NormalizedName, c.StudentId, c.SuggestedStudentId, suggested.DisplayName, suggestedSchool.Name, c.LinkStatus
                FROM dbo.EntryCompetitors c
                INNER JOIN dbo.Entries en ON en.Id = c.EntryId
                INNER JOIN dbo.Events e ON e.Id = en.EventId
                LEFT JOIN dbo.Students suggested ON suggested.Id = c.SuggestedStudentId
                LEFT JOIN dbo.Schools suggestedSchool ON suggestedSchool.Id = suggested.SchoolId
                WHERE e.TournamentId = @Id
                ORDER BY c.Id
                """, connection);
            command.Parameters.AddWithValue("@Id", id);
            await using var reader = await command.ExecuteReaderAsync(cancellationToken);
            while (await reader.ReadAsync(cancellationToken))
            {
                entries[reader.GetInt32(1)].Competitors.Add(new CompetitorRow
                {
                    Id = reader.GetInt32(0),
                    RawName = reader.GetString(2),
                    NormalizedName = reader.GetString(3),
                    StudentId = reader.IsDBNull(4) ? null : reader.GetInt32(4),
                    SuggestedStudentId = reader.IsDBNull(5) ? null : reader.GetInt32(5),
                    SuggestedName = reader.IsDBNull(6) ? null : reader.GetString(6),
                    SuggestedSchool = reader.IsDBNull(7) ? null : reader.GetString(7),
                    LinkStatus = reader.GetString(8)
                });
            }
        }

        tournament.Events.AddRange(events.Values);
        return tournament;
    }

    public async Task<IReadOnlyList<SchoolStudent>> StudentsForSchoolsAsync(IEnumerable<int> schoolIds, CancellationToken cancellationToken)
    {
        var ids = schoolIds.Distinct().ToArray();
        if (ids.Length == 0)
        {
            return Array.Empty<SchoolStudent>();
        }

        await using var connection = await OpenAsync(cancellationToken);
        await using var command = new SqlCommand(
            $"SELECT s.Id, s.SchoolId, s.DisplayName, sc.Name FROM dbo.Students s INNER JOIN dbo.Schools sc ON sc.Id = s.SchoolId WHERE s.SchoolId IN ({string.Join(",", ids)}) ORDER BY s.DisplayName",
            connection);
        var students = new List<SchoolStudent>();
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        while (await reader.ReadAsync(cancellationToken))
        {
            students.Add(new SchoolStudent
            {
                Id = reader.GetInt32(0),
                SchoolId = reader.GetInt32(1),
                DisplayName = reader.GetString(2),
                SchoolName = reader.GetString(3)
            });
        }

        return students;
    }

    public async Task<int> SaveTabroomImportAsync(
        TabroomTournamentInfo info,
        IReadOnlyList<ImportedEvent> events,
        string skipped,
        int? existingPreviewId,
        CancellationToken cancellationToken)
    {
        await using var connection = await OpenAsync(cancellationToken);
        await using var transaction = (SqlTransaction)await connection.BeginTransactionAsync(cancellationToken);
        var tournamentId = existingPreviewId ?? await InsertTournamentAsync(connection, transaction, info, skipped, cancellationToken);
        if (existingPreviewId.HasValue)
        {
            await DeleteEventsAsync(connection, transaction, existingPreviewId.Value, cancellationToken);
            await using var update = new SqlCommand("""
                UPDATE dbo.Tournaments
                SET Name = @Name, StartDate = @StartDate, StateCode = @State, InState = @InState,
                    Kind = @Kind, SkippedSummary = @Skipped, RawJson = @RawJson, Status = N'Preview'
                WHERE Id = @Id
                """, connection, transaction);
            update.Parameters.AddWithValue("@Name", info.Name);
            update.Parameters.AddWithValue("@StartDate", info.StartDate.ToDateTime(TimeOnly.MinValue));
            update.Parameters.AddWithValue("@State", (object?)info.State ?? DBNull.Value);
            update.Parameters.AddWithValue("@InState", info.InState);
            update.Parameters.AddWithValue("@Kind", info.SuggestedKind);
            update.Parameters.AddWithValue("@Skipped", string.IsNullOrWhiteSpace(skipped) ? DBNull.Value : skipped);
            update.Parameters.AddWithValue("@RawJson", info.RawJson);
            update.Parameters.AddWithValue("@Id", existingPreviewId.Value);
            await update.ExecuteNonQueryAsync(cancellationToken);
        }

        var known = await LoadKnownStudentsAsync(connection, transaction, cancellationToken);
        foreach (var imported in events)
        {
            await InsertEventAsync(connection, transaction, tournamentId, imported, known, cancellationToken);
        }

        await transaction.CommitAsync(cancellationToken);
        return tournamentId;
    }

    public async Task<int> CreateManualAsync(string name, DateOnly start, string? state, string kind, CancellationToken cancellationToken)
    {
        var stateCode = string.IsNullOrWhiteSpace(state) ? null : state.Trim();
        await using var connection = await OpenAsync(cancellationToken);
        await using var command = new SqlCommand("""
            INSERT INTO dbo.Tournaments (TabroomTournamentId, Name, StartDate, StateCode, InState, Kind, Status)
            OUTPUT INSERTED.Id
            VALUES (NULL, @Name, @StartDate, @State, @InState, @Kind, N'Preview')
            """, connection);
        command.Parameters.AddWithValue("@Name", name.Trim());
        command.Parameters.AddWithValue("@StartDate", start.ToDateTime(TimeOnly.MinValue));
        command.Parameters.AddWithValue("@State", (object?)stateCode ?? DBNull.Value);
        command.Parameters.AddWithValue("@InState", string.Equals(stateCode, "MN", StringComparison.OrdinalIgnoreCase));
        command.Parameters.AddWithValue("@Kind", kind);
        return (int)(await command.ExecuteScalarAsync(cancellationToken) ?? 0);
    }

    public async Task AddEntryAsync(
        int tournamentId,
        string eventKind,
        string eventName,
        int place,
        string schoolName,
        string competitorNames,
        bool prelim,
        int elimWins,
        bool reachedFinals,
        bool reachedSemis,
        bool qualified,
        CancellationToken cancellationToken)
    {
        await using var connection = await OpenAsync(cancellationToken);
        await using var transaction = (SqlTransaction)await connection.BeginTransactionAsync(cancellationToken);
        var eventId = await FindOrCreateEventAsync(connection, transaction, tournamentId, eventKind, eventName, cancellationToken);
        var schoolId = await FindOrCreateSchoolAsync(connection, transaction, null, schoolName, cancellationToken);
        var known = await LoadKnownStudentsAsync(connection, transaction, cancellationToken);
        var names = NameNormalizer.SplitCompetitors(competitorNames);
        var entryId = await InsertEntryRowAsync(connection, transaction, eventId, schoolId, competitorNames, null, place, null, prelim, !prelim, elimWins, reachedFinals, reachedSemis, qualified, cancellationToken);
        foreach (var name in names)
        {
            await InsertCompetitorAsync(connection, transaction, entryId, schoolId, name, known, cancellationToken);
        }

        await RecalculateEventCountsAsync(connection, transaction, eventId, cancellationToken);
        await transaction.CommitAsync(cancellationToken);
    }

    public async Task SaveReviewAsync(int tournamentId, ReviewUpdate update, CancellationToken cancellationToken)
    {
        await using var connection = await OpenAsync(cancellationToken);
        await using var transaction = (SqlTransaction)await connection.BeginTransactionAsync(cancellationToken);
        if (Enum.TryParse<TournamentKind>(update.Kind, out _))
        {
            await using var command = new SqlCommand("UPDATE dbo.Tournaments SET Kind = @Kind WHERE Id = @Id", connection, transaction);
            command.Parameters.AddWithValue("@Kind", update.Kind);
            command.Parameters.AddWithValue("@Id", tournamentId);
            await command.ExecuteNonQueryAsync(cancellationToken);
        }

        foreach (var (eventId, octos) in update.Octos)
        {
            await using var command = new SqlCommand("""
                UPDATE dbo.Events
                SET BreaksToOctosOrDeeper = @Octos, HasFinalSession = @Finals, HasSemis = @Semis
                WHERE Id = @Id AND TournamentId = @TournamentId
                """, connection, transaction);
            command.Parameters.AddWithValue("@Octos", octos);
            command.Parameters.AddWithValue("@Finals", update.FinalSessions.GetValueOrDefault(eventId));
            command.Parameters.AddWithValue("@Semis", update.Semis.GetValueOrDefault(eventId));
            command.Parameters.AddWithValue("@Id", eventId);
            command.Parameters.AddWithValue("@TournamentId", tournamentId);
            await command.ExecuteNonQueryAsync(cancellationToken);
        }

        foreach (var (entryId, entry) in update.Entries)
        {
            var correctPlace = entry.Place is >= 1;
            var note = string.IsNullOrWhiteSpace(entry.PlaceNote) ? null : entry.PlaceNote.Trim();
            if (note is { Length: > 400 })
            {
                note = note[..400];
            }
            await using var command = new SqlCommand("""
                UPDATE dbo.Entries
                SET ElimWins = @ElimWins, ReachedElims = @ReachedElims, ReachedFinals = @ReachedFinals,
                    ReachedSemis = @ReachedSemis, QualifiedToState = @Qualified,
                    Place = CASE WHEN @SetPlace = 1 THEN @Place ELSE Place END,
                    PlaceNote = CASE WHEN @SetPlace = 1 THEN @PlaceNote ELSE PlaceNote END
                WHERE Id = @Id AND EventId IN (SELECT Id FROM dbo.Events WHERE TournamentId = @TournamentId)
                """, connection, transaction);
            command.Parameters.AddWithValue("@ElimWins", Math.Max(0, entry.ElimWins));
            command.Parameters.AddWithValue("@ReachedElims", entry.ReachedElims);
            command.Parameters.AddWithValue("@ReachedFinals", entry.ReachedFinals);
            command.Parameters.AddWithValue("@ReachedSemis", entry.ReachedSemis);
            command.Parameters.AddWithValue("@Qualified", entry.QualifiedToState);
            command.Parameters.AddWithValue("@SetPlace", correctPlace ? 1 : 0);
            command.Parameters.AddWithValue("@Place", correctPlace ? entry.Place!.Value : 0);
            command.Parameters.AddWithValue("@PlaceNote", (object?)note ?? DBNull.Value);
            command.Parameters.AddWithValue("@Id", entryId);
            command.Parameters.AddWithValue("@TournamentId", tournamentId);
            await command.ExecuteNonQueryAsync(cancellationToken);
        }

        foreach (var (competitorId, choice) in update.CompetitorChoices)
        {
            int? studentId = null;
            var status = "Possible";
            if (string.Equals(choice, "new", StringComparison.OrdinalIgnoreCase))
            {
                status = "New";
            }
            else if (int.TryParse(choice, out var parsed))
            {
                status = "Linked";
                studentId = parsed;
            }

            await using var command = new SqlCommand("""
                UPDATE dbo.EntryCompetitors
                SET LinkStatus = @Status, StudentId = @StudentId
                WHERE Id = @Id AND EntryId IN (
                    SELECT en.Id FROM dbo.Entries en INNER JOIN dbo.Events e ON e.Id = en.EventId WHERE e.TournamentId = @TournamentId)
                """, connection, transaction);
            command.Parameters.AddWithValue("@Status", status);
            command.Parameters.AddWithValue("@StudentId", (object?)studentId ?? DBNull.Value);
            command.Parameters.AddWithValue("@Id", competitorId);
            command.Parameters.AddWithValue("@TournamentId", tournamentId);
            await command.ExecuteNonQueryAsync(cancellationToken);
        }

        await transaction.CommitAsync(cancellationToken);
    }

    public async Task PublishAsync(int tournamentId, CancellationToken cancellationToken)
    {
        var tournament = await GetTournamentAsync(tournamentId, cancellationToken)
            ?? throw new InvalidOperationException("Tournament was not found.");
        var statuses = tournament.Events.SelectMany(item => item.Entries).SelectMany(item => item.Competitors).Select(item => Enum.Parse<LinkStatus>(item.LinkStatus));
        if (!NameLinker.CanPublish(statuses))
        {
            throw new InvalidOperationException("Resolve every possible name match before publishing.");
        }

        await using var connection = await OpenAsync(cancellationToken);
        await using var transaction = (SqlTransaction)await connection.BeginTransactionAsync(cancellationToken);
        var created = new Dictionary<(int SchoolId, string Name), int>();
        foreach (var entry in tournament.Events.SelectMany(item => item.Entries))
        {
            foreach (var competitor in entry.Competitors)
            {
                var studentId = competitor.StudentId;
                if (competitor.LinkStatus == LinkStatus.New.ToString() || studentId is null)
                {
                    var key = (entry.SchoolId, competitor.NormalizedName);
                    if (!created.TryGetValue(key, out var existing))
                    {
                        existing = await InsertStudentAsync(connection, transaction, entry.SchoolId, competitor.RawName, competitor.NormalizedName, cancellationToken);
                        created[key] = existing;
                    }

                    studentId = existing;
                    await using var link = new SqlCommand(
                        "UPDATE dbo.EntryCompetitors SET StudentId = @StudentId, LinkStatus = N'Linked' WHERE Id = @Id",
                        connection,
                        transaction);
                    link.Parameters.AddWithValue("@StudentId", studentId);
                    link.Parameters.AddWithValue("@Id", competitor.Id);
                    await link.ExecuteNonQueryAsync(cancellationToken);
                }

                await EnsureAliasAsync(connection, transaction, studentId.Value, competitor.NormalizedName, cancellationToken);
            }
        }

        await using var publish = new SqlCommand("UPDATE dbo.Tournaments SET Status = N'Published' WHERE Id = @Id", connection, transaction);
        publish.Parameters.AddWithValue("@Id", tournamentId);
        await publish.ExecuteNonQueryAsync(cancellationToken);
        await transaction.CommitAsync(cancellationToken);
    }

    private static async Task<int> InsertTournamentAsync(SqlConnection connection, SqlTransaction transaction, TabroomTournamentInfo info, string skipped, CancellationToken cancellationToken)
    {
        await using var command = new SqlCommand("""
            INSERT INTO dbo.Tournaments (TabroomTournamentId, Name, StartDate, StateCode, InState, Kind, Status, SkippedSummary, RawJson)
            OUTPUT INSERTED.Id
            VALUES (@TabroomId, @Name, @StartDate, @State, @InState, @Kind, N'Preview', @Skipped, @RawJson)
            """, connection, transaction);
        command.Parameters.AddWithValue("@TabroomId", info.Id);
        command.Parameters.AddWithValue("@Name", info.Name);
        command.Parameters.AddWithValue("@StartDate", info.StartDate.ToDateTime(TimeOnly.MinValue));
        command.Parameters.AddWithValue("@State", (object?)info.State ?? DBNull.Value);
        command.Parameters.AddWithValue("@InState", info.InState);
        command.Parameters.AddWithValue("@Kind", info.SuggestedKind);
        command.Parameters.AddWithValue("@Skipped", string.IsNullOrWhiteSpace(skipped) ? DBNull.Value : skipped);
        command.Parameters.AddWithValue("@RawJson", info.RawJson);
        return (int)(await command.ExecuteScalarAsync(cancellationToken) ?? 0);
    }

    private static async Task DeleteEventsAsync(SqlConnection connection, SqlTransaction transaction, int tournamentId, CancellationToken cancellationToken)
    {
        foreach (var sql in new[]
        {
            """
            DELETE c FROM dbo.EntryCompetitors c
            INNER JOIN dbo.Entries en ON en.Id = c.EntryId
            INNER JOIN dbo.Events e ON e.Id = en.EventId
            WHERE e.TournamentId = @Id
            """,
            """
            DELETE en FROM dbo.Entries en
            INNER JOIN dbo.Events e ON e.Id = en.EventId
            WHERE e.TournamentId = @Id
            """,
            "DELETE FROM dbo.Events WHERE TournamentId = @Id"
        })
        {
            await using var command = new SqlCommand(sql, connection, transaction);
            command.Parameters.AddWithValue("@Id", tournamentId);
            await command.ExecuteNonQueryAsync(cancellationToken);
        }
    }

    private static async Task InsertEventAsync(
        SqlConnection connection,
        SqlTransaction transaction,
        int tournamentId,
        ImportedEvent imported,
        List<KnownStudent> known,
        CancellationToken cancellationToken)
    {
        await using var command = new SqlCommand("""
            INSERT INTO dbo.Events
                (TournamentId, TabroomEventId, TabroomResultSetId, Name, EventKind, CountedDivision, SchoolCount, EntryCount,
                 BreaksToOctosOrDeeper, HasFinalSession, HasSemis, IsHouse, RawJson)
            OUTPUT INSERTED.Id
            VALUES
                (@TournamentId, @EventId, @ResultId, @Name, @EventKind, @Counted, @Schools, @Entries,
                 @Octos, @Finals, @Semis, @House, @RawJson)
            """, connection, transaction);
        command.Parameters.AddWithValue("@TournamentId", tournamentId);
        command.Parameters.AddWithValue("@EventId", imported.TabroomEventId);
        command.Parameters.AddWithValue("@ResultId", (object?)imported.ResultSetId ?? DBNull.Value);
        command.Parameters.AddWithValue("@Name", imported.Name);
        command.Parameters.AddWithValue("@EventKind", imported.EventKind);
        command.Parameters.AddWithValue("@Counted", imported.CountedDivision);
        command.Parameters.AddWithValue("@Schools", imported.SchoolCount);
        command.Parameters.AddWithValue("@Entries", imported.EntryCount);
        command.Parameters.AddWithValue("@Octos", imported.BreaksToOctosOrDeeper);
        command.Parameters.AddWithValue("@Finals", imported.HasFinalSession);
        command.Parameters.AddWithValue("@Semis", imported.HasSemis);
        command.Parameters.AddWithValue("@House", imported.IsHouse);
        command.Parameters.AddWithValue("@RawJson", imported.RawJson);
        var eventId = (int)(await command.ExecuteScalarAsync(cancellationToken) ?? 0);

        foreach (var entry in imported.Entries)
        {
            var schoolId = await FindOrCreateSchoolAsync(connection, transaction, entry.TabroomSchoolId, entry.SchoolName, cancellationToken);
            var entryId = await InsertEntryRowAsync(
                connection,
                transaction,
                eventId,
                schoolId,
                entry.EntryName,
                entry.TabroomEntryId,
                entry.Place,
                entry.Place,
                entry.IsPrelim,
                entry.ReachedElims,
                entry.ElimWins,
                entry.ReachedFinals,
                entry.ReachedSemis,
                false,
                cancellationToken);
            foreach (var competitor in entry.Competitors)
            {
                await InsertCompetitorAsync(connection, transaction, entryId, schoolId, competitor.Name, known, cancellationToken);
            }
        }
    }

    private static async Task<int> FindOrCreateEventAsync(SqlConnection connection, SqlTransaction transaction, int tournamentId, string eventKind, string eventName, CancellationToken cancellationToken)
    {
        await using var find = new SqlCommand(
            "SELECT Id FROM dbo.Events WHERE TournamentId = @TournamentId AND EventKind = @EventKind",
            connection,
            transaction);
        find.Parameters.AddWithValue("@TournamentId", tournamentId);
        find.Parameters.AddWithValue("@EventKind", eventKind);
        var existing = await find.ExecuteScalarAsync(cancellationToken);
        if (existing is int id)
        {
            return id;
        }

        await using var insert = new SqlCommand("""
            INSERT INTO dbo.Events
                (TournamentId, Name, EventKind, CountedDivision, SchoolCount, EntryCount, BreaksToOctosOrDeeper, HasFinalSession, HasSemis, IsHouse)
            OUTPUT INSERTED.Id
            VALUES (@TournamentId, @Name, @EventKind, 1, 0, 0, 0, @Finals, 0, @House)
            """, connection, transaction);
        insert.Parameters.AddWithValue("@TournamentId", tournamentId);
        insert.Parameters.AddWithValue("@Name", eventName);
        insert.Parameters.AddWithValue("@EventKind", eventKind);
        insert.Parameters.AddWithValue("@Finals", eventKind == EventKind.Congress.ToString());
        insert.Parameters.AddWithValue("@House", EventClassifier.IsHouse(eventName));
        return (int)(await insert.ExecuteScalarAsync(cancellationToken) ?? 0);
    }

    private static async Task<int> FindOrCreateSchoolAsync(SqlConnection connection, SqlTransaction transaction, int? tabroomSchoolId, string schoolName, CancellationToken cancellationToken)
    {
        if (tabroomSchoolId.HasValue)
        {
            await using var byTabroom = new SqlCommand("SELECT Id FROM dbo.Schools WHERE TabroomSchoolId = @Id", connection, transaction);
            byTabroom.Parameters.AddWithValue("@Id", tabroomSchoolId.Value);
            var found = await byTabroom.ExecuteScalarAsync(cancellationToken);
            if (found is int id)
            {
                return id;
            }
        }

        var normalized = NameNormalizer.Normalize(schoolName);
        await using var byName = new SqlCommand("SELECT TOP 1 Id FROM dbo.Schools WHERE NormalizedName = @Name ORDER BY Id", connection, transaction);
        byName.Parameters.AddWithValue("@Name", normalized);
        var named = await byName.ExecuteScalarAsync(cancellationToken);
        if (named is int schoolId)
        {
            if (tabroomSchoolId.HasValue)
            {
                await using var attach = new SqlCommand(
                    "UPDATE dbo.Schools SET TabroomSchoolId = @TabroomId WHERE Id = @Id AND TabroomSchoolId IS NULL",
                    connection,
                    transaction);
                attach.Parameters.AddWithValue("@TabroomId", tabroomSchoolId.Value);
                attach.Parameters.AddWithValue("@Id", schoolId);
                await attach.ExecuteNonQueryAsync(cancellationToken);
            }

            return schoolId;
        }

        await using var insert = new SqlCommand("""
            INSERT INTO dbo.Schools (TabroomSchoolId, Name, NormalizedName)
            OUTPUT INSERTED.Id
            VALUES (@TabroomId, @Name, @Normalized)
            """, connection, transaction);
        insert.Parameters.AddWithValue("@TabroomId", (object?)tabroomSchoolId ?? DBNull.Value);
        insert.Parameters.AddWithValue("@Name", schoolName.Trim());
        insert.Parameters.AddWithValue("@Normalized", normalized);
        return (int)(await insert.ExecuteScalarAsync(cancellationToken) ?? 0);
    }

    private static async Task<int> InsertEntryRowAsync(
        SqlConnection connection,
        SqlTransaction transaction,
        int eventId,
        int schoolId,
        string entryName,
        int? tabroomEntryId,
        int place,
        int? tabroomPlace,
        bool prelim,
        bool reachedElims,
        int elimWins,
        bool reachedFinals,
        bool reachedSemis,
        bool qualified,
        CancellationToken cancellationToken)
    {
        await using var command = new SqlCommand("""
            INSERT INTO dbo.Entries
                (EventId, TabroomEntryId, EntryName, SchoolId, Place, TabroomPlace, IsPrelim, ReachedElims, ElimWins, ReachedFinals, ReachedSemis, QualifiedToState)
            OUTPUT INSERTED.Id
            VALUES (@EventId, @TabroomEntryId, @EntryName, @SchoolId, @Place, @TabroomPlace, @Prelim, @ReachedElims, @ElimWins, @ReachedFinals, @ReachedSemis, @Qualified)
            """, connection, transaction);
        command.Parameters.AddWithValue("@EventId", eventId);
        command.Parameters.AddWithValue("@TabroomEntryId", (object?)tabroomEntryId ?? DBNull.Value);
        command.Parameters.AddWithValue("@EntryName", entryName);
        command.Parameters.AddWithValue("@SchoolId", schoolId);
        command.Parameters.AddWithValue("@Place", place);
        command.Parameters.AddWithValue("@TabroomPlace", (object?)tabroomPlace ?? DBNull.Value);
        command.Parameters.AddWithValue("@Prelim", prelim);
        command.Parameters.AddWithValue("@ReachedElims", reachedElims);
        command.Parameters.AddWithValue("@ElimWins", elimWins);
        command.Parameters.AddWithValue("@ReachedFinals", reachedFinals);
        command.Parameters.AddWithValue("@ReachedSemis", reachedSemis);
        command.Parameters.AddWithValue("@Qualified", qualified);
        return (int)(await command.ExecuteScalarAsync(cancellationToken) ?? 0);
    }

    private static async Task InsertCompetitorAsync(
        SqlConnection connection,
        SqlTransaction transaction,
        int entryId,
        int schoolId,
        string rawName,
        List<KnownStudent> known,
        CancellationToken cancellationToken)
    {
        var decision = NameLinker.Classify(rawName, schoolId.ToString(), known);
        int? studentId = decision.Status == LinkStatus.Linked && int.TryParse(decision.StudentId, out var linked) ? linked : null;
        int? suggested = int.TryParse(decision.SuggestedStudentId, out var suggestion) ? suggestion : null;
        await using var command = new SqlCommand("""
            INSERT INTO dbo.EntryCompetitors (EntryId, RawName, NormalizedName, StudentId, SuggestedStudentId, LinkStatus)
            VALUES (@EntryId, @RawName, @Normalized, @StudentId, @Suggested, @Status)
            """, connection, transaction);
        command.Parameters.AddWithValue("@EntryId", entryId);
        command.Parameters.AddWithValue("@RawName", rawName.Trim());
        command.Parameters.AddWithValue("@Normalized", NameNormalizer.Normalize(rawName));
        command.Parameters.AddWithValue("@StudentId", (object?)studentId ?? DBNull.Value);
        command.Parameters.AddWithValue("@Suggested", (object?)suggested ?? DBNull.Value);
        command.Parameters.AddWithValue("@Status", decision.Status.ToString());
        await command.ExecuteNonQueryAsync(cancellationToken);
    }

    private static async Task<List<KnownStudent>> LoadKnownStudentsAsync(SqlConnection connection, SqlTransaction transaction, CancellationToken cancellationToken)
    {
        var students = new Dictionary<int, (int SchoolId, string Name, List<string> Aliases)>();
        await using (var command = new SqlCommand("SELECT Id, SchoolId, NormalizedName FROM dbo.Students", connection, transaction))
        await using (var reader = await command.ExecuteReaderAsync(cancellationToken))
        {
            while (await reader.ReadAsync(cancellationToken))
            {
                students.Add(reader.GetInt32(0), (reader.GetInt32(1), reader.GetString(2), new List<string>()));
            }
        }

        await using (var command = new SqlCommand("SELECT StudentId, NormalizedName FROM dbo.StudentAliases", connection, transaction))
        await using (var reader = await command.ExecuteReaderAsync(cancellationToken))
        {
            while (await reader.ReadAsync(cancellationToken))
            {
                if (students.TryGetValue(reader.GetInt32(0), out var student))
                {
                    student.Aliases.Add(reader.GetString(1));
                }
            }
        }

        return students.Select(pair => new KnownStudent(pair.Key.ToString(), pair.Value.SchoolId.ToString(), pair.Value.Name, pair.Value.Aliases)).ToList();
    }

    private static async Task<int> InsertStudentAsync(SqlConnection connection, SqlTransaction transaction, int schoolId, string displayName, string normalizedName, CancellationToken cancellationToken)
    {
        await using var command = new SqlCommand("""
            INSERT INTO dbo.Students (SchoolId, DisplayName, NormalizedName)
            OUTPUT INSERTED.Id
            VALUES (@SchoolId, @DisplayName, @Normalized)
            """, connection, transaction);
        command.Parameters.AddWithValue("@SchoolId", schoolId);
        command.Parameters.AddWithValue("@DisplayName", displayName.Trim());
        command.Parameters.AddWithValue("@Normalized", normalizedName);
        return (int)(await command.ExecuteScalarAsync(cancellationToken) ?? 0);
    }

    private static async Task EnsureAliasAsync(SqlConnection connection, SqlTransaction transaction, int studentId, string normalizedName, CancellationToken cancellationToken)
    {
        await using var command = new SqlCommand("""
            IF EXISTS (SELECT 1 FROM dbo.Students WHERE Id = @Id AND NormalizedName = @Name)
                RETURN;
            IF NOT EXISTS (SELECT 1 FROM dbo.StudentAliases WHERE StudentId = @Id AND NormalizedName = @Name)
                INSERT INTO dbo.StudentAliases (StudentId, NormalizedName) VALUES (@Id, @Name);
            """, connection, transaction);
        command.Parameters.AddWithValue("@Id", studentId);
        command.Parameters.AddWithValue("@Name", normalizedName);
        await command.ExecuteNonQueryAsync(cancellationToken);
    }

    private static async Task RecalculateEventCountsAsync(SqlConnection connection, SqlTransaction transaction, int eventId, CancellationToken cancellationToken)
    {
        await using var command = new SqlCommand("""
            UPDATE dbo.Events
            SET EntryCount = (SELECT COUNT(*) FROM dbo.Entries WHERE EventId = @Id),
                SchoolCount = (SELECT COUNT(DISTINCT SchoolId) FROM dbo.Entries WHERE EventId = @Id)
            WHERE Id = @Id
            """, connection, transaction);
        command.Parameters.AddWithValue("@Id", eventId);
        await command.ExecuteNonQueryAsync(cancellationToken);
    }
}
