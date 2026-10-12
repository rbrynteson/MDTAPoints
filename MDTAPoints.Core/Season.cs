using System.Globalization;

namespace MDTAPoints.Core;

public static class SeasonWeek
{
    public static string Key(DateOnly date)
    {
        var day = date.ToDateTime(TimeOnly.MinValue);
        return $"{ISOWeek.GetYear(day)}-W{ISOWeek.GetWeekOfYear(day):00}";
    }
}

public sealed record SeasonCompetitor(string StudentId, string Name, string SchoolId, string SchoolName, int? Grade);

public sealed record SeasonEntry(
    string TournamentId,
    string TournamentName,
    DateOnly StartDate,
    bool InState,
    TournamentKind TournamentKind,
    EventKind Event,
    bool VarsityOrOpen,
    int SchoolCount,
    int EntryCount,
    bool BreaksToOctosOrDeeper,
    bool HasFinalSession,
    bool HasSemis,
    bool IsHouse,
    string EntryId,
    string SchoolId,
    string SchoolName,
    int Place,
    bool ReachedElims,
    int ElimWins,
    bool ReachedFinals,
    bool ReachedSemis,
    IReadOnlyList<SeasonCompetitor> Competitors);

public sealed record ManualQualification(string StudentId, EventKind Event, bool Qualified);

public sealed record HeadToHeadResult(string WinnerStudentId, string LoserStudentId);

public sealed record SeasonInput(
    IReadOnlyList<SeasonEntry> Entries,
    IReadOnlyList<ManualQualification> Qualifications,
    IReadOnlyList<HeadToHeadResult> HeadToHead);

public sealed record PointLine(
    string TournamentId,
    string TournamentName,
    DateOnly Start,
    int Points,
    int Place,
    bool Dropped,
    string Detail);

public sealed record IndividualStanding(
    string StudentId,
    string Name,
    string SchoolId,
    string SchoolName,
    EventKind Event,
    int Points,
    bool Eligible,
    int TournamentCount,
    bool QualifiedToState,
    int? Grade,
    double AveragePlace,
    int? AwardRank,
    IReadOnlyList<PointLine> Lines);

public sealed record TeamStanding(
    string SchoolId,
    string SchoolName,
    EventKind Event,
    int Points,
    int Rank,
    int TournamentCount,
    double AveragePlace,
    IReadOnlyList<PointLine> Lines);

public sealed record SeasonStandings(
    IReadOnlyList<IndividualStanding> Individuals,
    IReadOnlyList<TeamStanding> Teams);
