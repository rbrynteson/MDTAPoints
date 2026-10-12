namespace MDTAPoints.Core;

public enum EventKind
{
    LincolnDouglas,
    Classic,
    PublicForum,
    Policy,
    Congress
}

public enum TournamentKind
{
    Invitational,
    Section,
    State,
    NsdaQualifier
}

public enum LinkStatus
{
    Linked,
    Possible,
    New
}

public sealed record ScoreRequest(
    TournamentKind TournamentKind,
    EventKind Event,
    bool VarsityOrOpen,
    bool InState,
    int SchoolCount,
    int EntryCount,
    bool BreaksToOctosOrDeeper,
    bool HasFinalSession,
    bool HasSemis,
    int Place,
    bool ReachedElims,
    int ElimWins,
    bool ReachedFinals,
    bool ReachedSemis,
    bool IsHouse);

public sealed record ScoreResult(int Points, bool Counts, IReadOnlyList<string> Reasons);

public static class EventKindLabels
{
    public static string Display(EventKind kind) => kind switch
    {
        EventKind.LincolnDouglas => "Lincoln-Douglas",
        EventKind.Classic => "Classic",
        EventKind.PublicForum => "Public Forum",
        EventKind.Policy => "Policy",
        EventKind.Congress => "Congressional Debate",
        _ => kind.ToString()
    };

    public static bool IsTeamEvent(EventKind kind) =>
        kind is EventKind.PublicForum or EventKind.Policy or EventKind.Classic;
}
