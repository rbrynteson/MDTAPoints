namespace MDTAPoints.Core;

public sealed record ParticipationCompetitor(int? StudentId, string NormalizedName);

public sealed record ParticipationEntry(
    EventKind Event,
    int EntryId,
    int SchoolId,
    string SchoolName,
    IReadOnlyList<ParticipationCompetitor> Competitors);

public sealed record EventParticipation(EventKind Event, int SchoolCount, int Count, bool CountEntries);

public sealed record SchoolParticipation(int SchoolId, string SchoolName, IReadOnlyDictionary<EventKind, int> Counts, int Total);

public sealed record ParticipationReport(
    IReadOnlyList<EventParticipation> Events,
    IReadOnlyList<SchoolParticipation> Schools,
    int StudentCount);

public static class ParticipationAnalytics
{
    public static readonly EventKind[] EventOrder =
    {
        EventKind.LincolnDouglas,
        EventKind.Classic,
        EventKind.PublicForum,
        EventKind.Policy,
        EventKind.Congress
    };

    public static ParticipationReport Calculate(IEnumerable<ParticipationEntry> source)
    {
        var entries = source.ToList();
        var events = EventOrder
            .Select(kind => SummarizeEvent(kind, entries.Where(entry => entry.Event == kind).ToList()))
            .ToList();
        var schools = entries
            .GroupBy(entry => entry.SchoolId)
            .Select(group => SummarizeSchool(group.Key, group.First().SchoolName, group.ToList()))
            .OrderBy(school => school.SchoolName, StringComparer.OrdinalIgnoreCase)
            .ThenBy(school => school.SchoolId)
            .ToList();
        return new ParticipationReport(events, schools, People(entries).Count);
    }

    public static string SchoolsAndStudents(int schools, int students)
    {
        var schoolLabel = schools == 1 ? "school" : "schools";
        var studentLabel = students == 1 ? "student" : "students";
        return $"{schools} {schoolLabel}, {students} {studentLabel}";
    }

    public static string CountLabel(bool countEntries) => countEntries ? "Entries" : "Students";

    private static EventParticipation SummarizeEvent(EventKind kind, List<ParticipationEntry> entries)
    {
        var countEntries = EventKindLabels.IsTeamEvent(kind);
        var count = countEntries
            ? entries.Select(PartnershipKey).Distinct(StringComparer.Ordinal).Count()
            : People(entries).Count;
        var schools = entries.Select(entry => entry.SchoolId).Distinct().Count();
        return new EventParticipation(kind, schools, count, countEntries);
    }

    private static SchoolParticipation SummarizeSchool(int schoolId, string schoolName, List<ParticipationEntry> entries)
    {
        var counts = new Dictionary<EventKind, int>();
        foreach (var kind in EventOrder)
        {
            var inEvent = entries.Where(entry => entry.Event == kind).ToList();
            counts[kind] = EventKindLabels.IsTeamEvent(kind)
                ? inEvent.Select(PartnershipKey).Distinct(StringComparer.Ordinal).Count()
                : People(inEvent).Count;
        }

        return new SchoolParticipation(schoolId, schoolName, counts, counts.Values.Sum());
    }

    private static HashSet<string> People(IReadOnlyList<ParticipationEntry> entries)
    {
        var linkedNames = new HashSet<string>(StringComparer.Ordinal);
        var people = new List<(bool Linked, string Key, string? NameKey)>();
        foreach (var entry in entries)
        {
            var added = false;
            foreach (var competitor in entry.Competitors)
            {
                var token = CompetitorToken(competitor);
                if (token.Length == 0)
                {
                    continue;
                }

                added = true;
                if (competitor.StudentId is int)
                {
                    var name = NameKey(entry.SchoolId, competitor.NormalizedName);
                    if (name is not null)
                    {
                        linkedNames.Add(name);
                    }

                    people.Add((true, "id:" + competitor.StudentId.Value, null));
                }
                else
                {
                    var name = NameKey(entry.SchoolId, competitor.NormalizedName);
                    people.Add((false, "name:" + name, name));
                }
            }

            if (!added)
            {
                people.Add((true, "entry:" + entry.EntryId, null));
            }
        }

        var set = new HashSet<string>(StringComparer.Ordinal);
        foreach (var person in people)
        {
            if (!person.Linked && person.NameKey is not null && linkedNames.Contains(person.NameKey))
            {
                continue;
            }

            set.Add(person.Key);
        }

        return set;
    }

    private static string PartnershipKey(ParticipationEntry entry)
    {
        var parts = entry.Competitors
            .Select(CompetitorToken)
            .Where(token => token.Length > 0)
            .Distinct(StringComparer.Ordinal)
            .OrderBy(token => token, StringComparer.Ordinal)
            .ToArray();
        if (parts.Length == 0)
        {
            return "entry:" + entry.EntryId;
        }

        return entry.SchoolId + "|" + string.Join("|", parts);
    }

    private static string CompetitorToken(ParticipationCompetitor competitor)
    {
        if (competitor.StudentId is int id)
        {
            return "id:" + id;
        }

        var name = competitor.NormalizedName.Trim().ToLowerInvariant();
        return name.Length == 0 ? "" : "name:" + name;
    }

    private static string? NameKey(int schoolId, string normalizedName)
    {
        var name = normalizedName.Trim().ToLowerInvariant();
        return name.Length == 0 ? null : schoolId + ":" + name;
    }
}
