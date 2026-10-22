namespace MDTAPoints.Api.Data;

public sealed class TournamentListItem
{
    public int Id { get; init; }
    public int? TabroomTournamentId { get; init; }
    public string Name { get; init; } = "";
    public DateOnly StartDate { get; init; }
    public string Kind { get; init; } = "";
    public string Status { get; init; } = "";
    public bool InState { get; init; }
}

public sealed class CompetitorRow
{
    public int Id { get; init; }
    public string RawName { get; init; } = "";
    public string NormalizedName { get; init; } = "";
    public int? StudentId { get; init; }
    public int? SuggestedStudentId { get; init; }
    public string? SuggestedName { get; init; }
    public string? SuggestedSchool { get; init; }
    public string LinkStatus { get; init; } = "";
}

public sealed class EntryRow
{
    public int Id { get; init; }
    public string EntryName { get; init; } = "";
    public int SchoolId { get; init; }
    public string SchoolName { get; init; } = "";
    public int Place { get; init; }
    public int? TabroomPlace { get; init; }
    public string? PlaceNote { get; init; }
    public bool IsPrelim { get; init; }
    public bool ReachedElims { get; init; }
    public int ElimWins { get; init; }
    public bool ReachedFinals { get; init; }
    public bool ReachedSemis { get; init; }
    public bool QualifiedToState { get; init; }
    public List<CompetitorRow> Competitors { get; init; } = new();
}

public sealed class EventRow
{
    public int Id { get; init; }
    public string Name { get; init; } = "";
    public string EventKind { get; init; } = "";
    public int SchoolCount { get; init; }
    public int EntryCount { get; init; }
    public bool BreaksToOctosOrDeeper { get; init; }
    public bool HasFinalSession { get; init; }
    public bool HasSemis { get; init; }
    public bool IsHouse { get; init; }
    public bool CountedDivision { get; init; }
    public List<EntryRow> Entries { get; init; } = new();
}

public sealed record TournamentResultRow(int Place, string Name, string SchoolName, int Points, string Scoring, string? PlaceNote);

public sealed class TournamentDetail
{
    public int Id { get; init; }
    public int? TabroomTournamentId { get; init; }
    public string Name { get; init; } = "";
    public DateOnly StartDate { get; init; }
    public string? StateCode { get; init; }
    public bool InState { get; init; }
    public string Kind { get; init; } = "";
    public string Status { get; init; } = "";
    public string? SkippedSummary { get; init; }
    public List<EventRow> Events { get; init; } = new();
}

public sealed class SchoolStudent
{
    public int Id { get; init; }
    public int SchoolId { get; init; }
    public string DisplayName { get; init; } = "";
    public string SchoolName { get; init; } = "";
}

public sealed class AppUser
{
    public int Id { get; init; }
    public string Username { get; init; } = "";
    public string DisplayName { get; init; } = "";
    public string PasswordHash { get; init; } = "";
}

public sealed class StudentAdminRow
{
    public int Id { get; init; }
    public string DisplayName { get; init; } = "";
    public string SchoolName { get; init; } = "";
    public int? Grade { get; init; }
    public HashSet<string> ManualEvents { get; init; } = new();
    public HashSet<string> SectionEvents { get; init; } = new();
}

public sealed class HeadToHeadRow
{
    public int WinnerStudentId { get; init; }
    public int LoserStudentId { get; init; }
    public string WinnerName { get; init; } = "";
    public string LoserName { get; init; } = "";
}

public sealed record ImportOutcome(int TournamentId, bool AlreadyPublished);
