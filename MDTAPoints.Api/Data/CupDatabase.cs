using MDTAPoints.Core;
using Microsoft.AspNetCore.Identity;
using Microsoft.Data.SqlClient;

namespace MDTAPoints.Api.Data;

public sealed partial class CupDatabase
{
    private static readonly PasswordHasher<string> Passwords = new();
    private readonly string _connectionString;

    public CupDatabase(IConfiguration configuration)
    {
        _connectionString = configuration.GetConnectionString("Cup")
            ?? throw new InvalidOperationException("Set ConnectionStrings:Cup to the Azure SQL database.");
    }

    public async Task SeedAdminAsync(IConfiguration configuration, CancellationToken cancellationToken)
    {
        var username = configuration["Admin:Username"];
        var password = configuration["Admin:Password"];
        if (string.IsNullOrWhiteSpace(username) || string.IsNullOrWhiteSpace(password))
        {
            return;
        }

        await using var connection = await OpenAsync(cancellationToken);
        await using var count = new SqlCommand("SELECT COUNT(*) FROM dbo.Users", connection);
        var existing = (int)(await count.ExecuteScalarAsync(cancellationToken) ?? 0);
        if (existing > 0)
        {
            return;
        }

        await InsertUserAsync(connection, null, username.Trim(), username.Trim(), password, cancellationToken);
    }

    public async Task<AppUser?> AuthenticateAsync(string username, string password, CancellationToken cancellationToken)
    {
        await using var connection = await OpenAsync(cancellationToken);
        var user = await FindUserAsync(connection, username, cancellationToken);
        if (user is null)
        {
            return null;
        }

        var result = Passwords.VerifyHashedPassword(user.Username, user.PasswordHash, password);
        return result == PasswordVerificationResult.Failed ? null : user;
    }

    public async Task<IReadOnlyList<AppUser>> ListUsersAsync(CancellationToken cancellationToken)
    {
        await using var connection = await OpenAsync(cancellationToken);
        var users = new List<AppUser>();
        await using var command = new SqlCommand("SELECT Id, Username, DisplayName, PasswordHash FROM dbo.Users ORDER BY Username", connection);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        while (await reader.ReadAsync(cancellationToken))
        {
            users.Add(ReadUser(reader));
        }

        return users;
    }

    public async Task AddUserAsync(string username, string displayName, string password, CancellationToken cancellationToken)
    {
        await using var connection = await OpenAsync(cancellationToken);
        await InsertUserAsync(connection, null, username.Trim(), displayName.Trim(), password, cancellationToken);
    }

    public async Task<bool> DeleteUserAsync(int id, int currentUserId, CancellationToken cancellationToken)
    {
        if (id == currentUserId)
        {
            return false;
        }

        await using var connection = await OpenAsync(cancellationToken);
        await using var command = new SqlCommand("DELETE FROM dbo.Users WHERE Id = @Id", connection);
        command.Parameters.AddWithValue("@Id", id);
        return await command.ExecuteNonQueryAsync(cancellationToken) == 1;
    }

    public async Task<SeasonInput> LoadSeasonAsync(CancellationToken cancellationToken)
    {
        await using var connection = await OpenAsync(cancellationToken);
        var entries = await LoadEntriesAsync(connection, cancellationToken);
        var qualifications = await LoadQualificationsAsync(connection, cancellationToken);
        var headToHead = await LoadHeadToHeadAsync(connection, cancellationToken);
        return new SeasonInput(entries, qualifications, headToHead);
    }

    public async Task<IReadOnlyList<StudentAdminRow>> ListStudentsAsync(CancellationToken cancellationToken)
    {
        await using var connection = await OpenAsync(cancellationToken);
        var students = new List<StudentAdminRow>();
        const string sql = """
            SELECT s.Id, s.DisplayName, sc.Name, s.Grade
            FROM dbo.Students s
            INNER JOIN dbo.Schools sc ON sc.Id = s.SchoolId
            ORDER BY sc.Name, s.DisplayName
            """;
        await using (var command = new SqlCommand(sql, connection))
        await using (var reader = await command.ExecuteReaderAsync(cancellationToken))
        {
            while (await reader.ReadAsync(cancellationToken))
            {
                students.Add(new StudentAdminRow
                {
                    Id = reader.GetInt32(0),
                    DisplayName = reader.GetString(1),
                    SchoolName = reader.GetString(2),
                    Grade = reader.IsDBNull(3) ? null : reader.GetInt32(3)
                });
            }
        }

        await FillQualificationsAsync(connection, students, cancellationToken);
        return students;
    }

    public async Task UpdateStudentAsync(int id, int? grade, IReadOnlyCollection<string> manualEvents, CancellationToken cancellationToken)
    {
        await using var connection = await OpenAsync(cancellationToken);
        await using var transaction = (SqlTransaction)await connection.BeginTransactionAsync(cancellationToken);
        await using (var command = new SqlCommand("UPDATE dbo.Students SET Grade = @Grade WHERE Id = @Id", connection, transaction))
        {
            command.Parameters.AddWithValue("@Id", id);
            command.Parameters.AddWithValue("@Grade", (object?)grade ?? DBNull.Value);
            await command.ExecuteNonQueryAsync(cancellationToken);
        }

        await using (var clear = new SqlCommand("DELETE FROM dbo.StudentQualifications WHERE StudentId = @Id", connection, transaction))
        {
            clear.Parameters.AddWithValue("@Id", id);
            await clear.ExecuteNonQueryAsync(cancellationToken);
        }

        foreach (var eventKind in manualEvents.Distinct(StringComparer.Ordinal))
        {
            await using var insert = new SqlCommand(
                "INSERT INTO dbo.StudentQualifications (StudentId, EventKind, QualifiedToState) VALUES (@Id, @EventKind, 1)",
                connection,
                transaction);
            insert.Parameters.AddWithValue("@Id", id);
            insert.Parameters.AddWithValue("@EventKind", eventKind);
            await insert.ExecuteNonQueryAsync(cancellationToken);
        }

        await transaction.CommitAsync(cancellationToken);
    }

    public async Task<IReadOnlyList<HeadToHeadRow>> ListHeadToHeadAsync(CancellationToken cancellationToken)
    {
        const string sql = """
            SELECT h.WinnerStudentId, h.LoserStudentId, w.DisplayName, l.DisplayName
            FROM dbo.HeadToHead h
            INNER JOIN dbo.Students w ON w.Id = h.WinnerStudentId
            INNER JOIN dbo.Students l ON l.Id = h.LoserStudentId
            ORDER BY w.DisplayName, l.DisplayName
            """;
        await using var connection = await OpenAsync(cancellationToken);
        var rows = new List<HeadToHeadRow>();
        await using var command = new SqlCommand(sql, connection);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        while (await reader.ReadAsync(cancellationToken))
        {
            rows.Add(new HeadToHeadRow
            {
                WinnerStudentId = reader.GetInt32(0),
                LoserStudentId = reader.GetInt32(1),
                WinnerName = reader.GetString(2),
                LoserName = reader.GetString(3)
            });
        }

        return rows;
    }

    public async Task AddHeadToHeadAsync(int winnerId, int loserId, CancellationToken cancellationToken)
    {
        if (winnerId == loserId)
        {
            return;
        }

        const string sql = """
            IF NOT EXISTS (SELECT 1 FROM dbo.HeadToHead WHERE WinnerStudentId = @Winner AND LoserStudentId = @Loser)
            INSERT INTO dbo.HeadToHead (WinnerStudentId, LoserStudentId) VALUES (@Winner, @Loser)
            """;
        await using var connection = await OpenAsync(cancellationToken);
        await using var command = new SqlCommand(sql, connection);
        command.Parameters.AddWithValue("@Winner", winnerId);
        command.Parameters.AddWithValue("@Loser", loserId);
        await command.ExecuteNonQueryAsync(cancellationToken);
    }

    public async Task DeleteHeadToHeadAsync(int winnerId, int loserId, CancellationToken cancellationToken)
    {
        await using var connection = await OpenAsync(cancellationToken);
        await using var command = new SqlCommand(
            "DELETE FROM dbo.HeadToHead WHERE WinnerStudentId = @Winner AND LoserStudentId = @Loser",
            connection);
        command.Parameters.AddWithValue("@Winner", winnerId);
        command.Parameters.AddWithValue("@Loser", loserId);
        await command.ExecuteNonQueryAsync(cancellationToken);
    }

    private async Task<List<SeasonEntry>> LoadEntriesAsync(SqlConnection connection, CancellationToken cancellationToken)
    {
        const string sql = """
            SELECT t.Id, t.Name, t.StartDate, t.InState, t.Kind,
                   e.EventKind, e.CountedDivision, e.SchoolCount, e.EntryCount, e.BreaksToOctosOrDeeper,
                   e.HasFinalSession, e.HasSemis, e.IsHouse,
                   en.Id, en.Place, en.ReachedElims, en.ElimWins, en.ReachedFinals, en.ReachedSemis,
                   s.Id, s.Name,
                   c.StudentId, st.DisplayName, st.Grade, ss.Id, ss.Name
            FROM dbo.Tournaments t
            INNER JOIN dbo.Events e ON e.TournamentId = t.Id
            INNER JOIN dbo.Entries en ON en.EventId = e.Id
            INNER JOIN dbo.Schools s ON s.Id = en.SchoolId
            INNER JOIN dbo.EntryCompetitors c ON c.EntryId = en.Id
            INNER JOIN dbo.Students st ON st.Id = c.StudentId
            INNER JOIN dbo.Schools ss ON ss.Id = st.SchoolId
            WHERE t.Status = N'Published' AND e.CountedDivision = 1
            ORDER BY en.Id, c.Id
            """;
        var grouped = new Dictionary<int, MutableEntry>();
        await using var command = new SqlCommand(sql, connection);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        while (await reader.ReadAsync(cancellationToken))
        {
            var entryId = reader.GetInt32(13);
            if (!grouped.TryGetValue(entryId, out var entry))
            {
                entry = new MutableEntry(reader);
                grouped.Add(entryId, entry);
            }

            entry.Competitors.Add(new SeasonCompetitor(
                reader.GetInt32(21).ToString(),
                reader.GetString(22),
                reader.GetInt32(24).ToString(),
                reader.GetString(25),
                reader.IsDBNull(23) ? null : reader.GetInt32(23)));
        }

        return grouped.Values.Select(entry => entry.ToSeason()).ToList();
    }

    private static async Task<List<ManualQualification>> LoadQualificationsAsync(SqlConnection connection, CancellationToken cancellationToken)
    {
        const string sql = """
            SELECT c.StudentId, e.EventKind
            FROM dbo.Entries en
            INNER JOIN dbo.Events e ON e.Id = en.EventId
            INNER JOIN dbo.Tournaments t ON t.Id = e.TournamentId
            INNER JOIN dbo.EntryCompetitors c ON c.EntryId = en.Id
            WHERE t.Status = N'Published' AND t.Kind = N'Section' AND en.QualifiedToState = 1 AND c.StudentId IS NOT NULL
            UNION
            SELECT StudentId, EventKind FROM dbo.StudentQualifications WHERE QualifiedToState = 1
            """;
        var results = new List<ManualQualification>();
        await using var command = new SqlCommand(sql, connection);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        while (await reader.ReadAsync(cancellationToken))
        {
            if (Enum.TryParse<EventKind>(reader.GetString(1), out var eventKind))
            {
                results.Add(new ManualQualification(reader.GetInt32(0).ToString(), eventKind, true));
            }
        }

        return results;
    }

    private static async Task<List<HeadToHeadResult>> LoadHeadToHeadAsync(SqlConnection connection, CancellationToken cancellationToken)
    {
        var results = new List<HeadToHeadResult>();
        await using var command = new SqlCommand("SELECT WinnerStudentId, LoserStudentId FROM dbo.HeadToHead", connection);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        while (await reader.ReadAsync(cancellationToken))
        {
            results.Add(new HeadToHeadResult(reader.GetInt32(0).ToString(), reader.GetInt32(1).ToString()));
        }

        return results;
    }

    private static async Task FillQualificationsAsync(SqlConnection connection, List<StudentAdminRow> students, CancellationToken cancellationToken)
    {
        var byId = students.ToDictionary(student => student.Id);
        const string manual = "SELECT StudentId, EventKind FROM dbo.StudentQualifications WHERE QualifiedToState = 1";
        await using (var command = new SqlCommand(manual, connection))
        await using (var reader = await command.ExecuteReaderAsync(cancellationToken))
        {
            while (await reader.ReadAsync(cancellationToken))
            {
                if (byId.TryGetValue(reader.GetInt32(0), out var student))
                {
                    student.ManualEvents.Add(reader.GetString(1));
                }
            }
        }

        const string section = """
            SELECT c.StudentId, e.EventKind
            FROM dbo.Entries en
            INNER JOIN dbo.Events e ON e.Id = en.EventId
            INNER JOIN dbo.Tournaments t ON t.Id = e.TournamentId
            INNER JOIN dbo.EntryCompetitors c ON c.EntryId = en.Id
            WHERE t.Status = N'Published' AND t.Kind = N'Section' AND en.QualifiedToState = 1 AND c.StudentId IS NOT NULL
            """;
        await using var sectionCommand = new SqlCommand(section, connection);
        await using var sectionReader = await sectionCommand.ExecuteReaderAsync(cancellationToken);
        while (await sectionReader.ReadAsync(cancellationToken))
        {
            if (byId.TryGetValue(sectionReader.GetInt32(0), out var student))
            {
                student.SectionEvents.Add(sectionReader.GetString(1));
            }
        }
    }

    private static async Task InsertUserAsync(SqlConnection connection, SqlTransaction? transaction, string username, string displayName, string password, CancellationToken cancellationToken)
    {
        await using var command = new SqlCommand(
            "INSERT INTO dbo.Users (Username, DisplayName, PasswordHash) VALUES (@Username, @DisplayName, @PasswordHash)",
            connection,
            transaction);
        command.Parameters.AddWithValue("@Username", username);
        command.Parameters.AddWithValue("@DisplayName", displayName);
        command.Parameters.AddWithValue("@PasswordHash", Passwords.HashPassword(username, password));
        await command.ExecuteNonQueryAsync(cancellationToken);
    }

    private static async Task<AppUser?> FindUserAsync(SqlConnection connection, string username, CancellationToken cancellationToken)
    {
        await using var command = new SqlCommand(
            "SELECT Id, Username, DisplayName, PasswordHash FROM dbo.Users WHERE Username = @Username",
            connection);
        command.Parameters.AddWithValue("@Username", username.Trim());
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        return await reader.ReadAsync(cancellationToken) ? ReadUser(reader) : null;
    }

    private static AppUser ReadUser(SqlDataReader reader) => new()
    {
        Id = reader.GetInt32(0),
        Username = reader.GetString(1),
        DisplayName = reader.GetString(2),
        PasswordHash = reader.GetString(3)
    };

    private async Task<SqlConnection> OpenAsync(CancellationToken cancellationToken)
    {
        var connection = new SqlConnection(_connectionString);
        await connection.OpenAsync(cancellationToken);
        return connection;
    }

    private sealed class MutableEntry
    {
        public MutableEntry(SqlDataReader reader)
        {
            TournamentId = reader.GetInt32(0).ToString();
            TournamentName = reader.GetString(1);
            StartDate = DateOnly.FromDateTime(reader.GetDateTime(2));
            InState = reader.GetBoolean(3);
            Kind = Enum.Parse<TournamentKind>(reader.GetString(4));
            Event = Enum.Parse<EventKind>(reader.GetString(5));
            VarsityOrOpen = reader.GetBoolean(6);
            SchoolCount = reader.GetInt32(7);
            EntryCount = reader.GetInt32(8);
            BreaksToOctos = reader.GetBoolean(9);
            HasFinalSession = reader.GetBoolean(10);
            HasSemis = reader.GetBoolean(11);
            IsHouse = reader.GetBoolean(12);
            EntryId = reader.GetInt32(13).ToString();
            Place = reader.GetInt32(14);
            ReachedElims = reader.GetBoolean(15);
            ElimWins = reader.GetInt32(16);
            ReachedFinals = reader.GetBoolean(17);
            ReachedSemis = reader.GetBoolean(18);
            SchoolId = reader.GetInt32(19).ToString();
            SchoolName = reader.GetString(20);
        }

        public string TournamentId { get; }
        public string TournamentName { get; }
        public DateOnly StartDate { get; }
        public bool InState { get; }
        public TournamentKind Kind { get; }
        public EventKind Event { get; }
        public bool VarsityOrOpen { get; }
        public int SchoolCount { get; }
        public int EntryCount { get; }
        public bool BreaksToOctos { get; }
        public bool HasFinalSession { get; }
        public bool HasSemis { get; }
        public bool IsHouse { get; }
        public string EntryId { get; }
        public int Place { get; }
        public bool ReachedElims { get; }
        public int ElimWins { get; }
        public bool ReachedFinals { get; }
        public bool ReachedSemis { get; }
        public string SchoolId { get; }
        public string SchoolName { get; }
        public List<SeasonCompetitor> Competitors { get; } = new();

        public SeasonEntry ToSeason() => new(
            TournamentId,
            TournamentName,
            StartDate,
            InState,
            Kind,
            Event,
            VarsityOrOpen,
            SchoolCount,
            EntryCount,
            BreaksToOctos,
            HasFinalSession,
            HasSemis,
            IsHouse,
            EntryId,
            SchoolId,
            SchoolName,
            Place,
            ReachedElims,
            ElimWins,
            ReachedFinals,
            ReachedSemis,
            Competitors);
    }
}
