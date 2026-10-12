namespace MDTAPoints.Tabroom;

public sealed record TabroomTournamentInfo(
    int Id,
    string Name,
    DateOnly StartDate,
    string? State,
    bool InState,
    string SuggestedKind,
    string RawJson);

public sealed record TabroomEventSummary(
    int EventId,
    string Name,
    string? Abbr,
    string? Level,
    string? Type,
    int? NsdaCategory,
    int? ResultSetId,
    bool ShouldImport);

public sealed record ImportedCompetitor(string Name);

public sealed record ImportedEntry(
    int? TabroomEntryId,
    string EntryName,
    int? TabroomSchoolId,
    string SchoolName,
    int Place,
    bool IsPrelim,
    bool ReachedElims,
    int ElimWins,
    bool ReachedFinals,
    bool ReachedSemis,
    IReadOnlyList<ImportedCompetitor> Competitors);

public sealed record ImportedEvent(
    int TabroomEventId,
    int? ResultSetId,
    string Name,
    string EventKind,
    bool CountedDivision,
    int SchoolCount,
    int EntryCount,
    bool BreaksToOctosOrDeeper,
    bool HasFinalSession,
    bool HasSemis,
    bool IsHouse,
    string RawJson,
    IReadOnlyList<ImportedEntry> Entries);
