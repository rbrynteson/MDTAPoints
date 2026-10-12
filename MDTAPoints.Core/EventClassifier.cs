namespace MDTAPoints.Core;

public static class EventClassifier
{
    public static EventKind? Classify(string eventName, int? nsdaCategory, string? type)
    {
        var name = eventName.ToLowerInvariant();
        var eventType = type?.ToLowerInvariant() ?? "";

        if (name.Contains("congress") || eventType == "congress" || nsdaCategory == 301)
        {
            return EventKind.Congress;
        }

        if (name.Contains("classic"))
        {
            return EventKind.Classic;
        }

        if (name.Contains("public forum") || name.Contains(" pf") || name.StartsWith("pf") || nsdaCategory == 104)
        {
            return EventKind.PublicForum;
        }

        if (name.Contains("lincoln") || name.Contains(" ld") || name.StartsWith("ld") || nsdaCategory == 102)
        {
            return EventKind.LincolnDouglas;
        }

        if (name.Contains("policy") || name.Contains("cross-ex") || name.Contains("cross ex") || name.Contains(" cx") || nsdaCategory is 101 or 103)
        {
            return EventKind.Policy;
        }

        return null;
    }

    public static bool IsCountedDivision(string eventName)
    {
        var name = eventName.ToLowerInvariant();
        if (name.Contains("novice") || name.Contains("junior varsity"))
        {
            return false;
        }

        var tokens = name.Split(new[] { ' ', '(', ')', '/', '-' }, StringSplitOptions.RemoveEmptyEntries);
        return !tokens.Any(token => token is "jv");
    }

    public static bool ShouldImport(string eventName, int? nsdaCategory, string? type) =>
        Classify(eventName, nsdaCategory, type) is not null && IsCountedDivision(eventName);

    public static bool IsHouse(string eventName) =>
        eventName.Contains("house", StringComparison.OrdinalIgnoreCase);

    public static TournamentKind SuggestKind(string tournamentName)
    {
        var words = Words(tournamentName);
        if (words.Contains("section"))
        {
            return TournamentKind.Section;
        }

        if (words.Contains("festival") || words.Contains("mshsl") || words.Contains("state"))
        {
            return TournamentKind.State;
        }

        if (words.Contains("nsda") || words.Contains("nfl") || words.Contains("district"))
        {
            return TournamentKind.NsdaQualifier;
        }

        return TournamentKind.Invitational;
    }

    private static HashSet<string> Words(string text) =>
        text.ToLowerInvariant()
            .Split(new[] { ' ', '(', ')', '/', '-', ',' }, StringSplitOptions.RemoveEmptyEntries)
            .ToHashSet(StringComparer.Ordinal);
}
