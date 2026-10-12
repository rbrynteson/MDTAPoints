using System.Text.Json;
using System.Text.Json.Serialization;
using MDTAPoints.Core;

namespace MDTAPoints.Tabroom;

public static class TabroomParser
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true,
        NumberHandling = JsonNumberHandling.AllowReadingFromString
    };

    public static TabroomTournamentInfo ParseTournament(string json)
    {
        var dto = JsonSerializer.Deserialize<TournamentDto>(json, JsonOptions)
            ?? throw new InvalidDataException("Tabroom tournament response was empty.");
        var start = LocalDate(dto.Start, dto.Tz);
        var state = dto.State?.Trim();
        return new TabroomTournamentInfo(
            dto.Id,
            dto.Name ?? "Tournament",
            start,
            state,
            string.Equals(state, "MN", StringComparison.OrdinalIgnoreCase),
            EventClassifier.SuggestKind(dto.Name ?? "").ToString(),
            json);
    }

    public static IReadOnlyList<TabroomEventSummary> ParseEvents(string json)
    {
        var events = JsonSerializer.Deserialize<Dictionary<string, EventDto>>(json, JsonOptions)
            ?? new Dictionary<string, EventDto>();
        return events.Values.Select(item =>
        {
            var resultId = ChooseResultSet(item.ResultSets);
            return new TabroomEventSummary(
                item.Id,
                item.Name ?? "",
                item.Abbr,
                item.Level,
                item.Type,
                item.NsdaCategory,
                resultId,
                EventClassifier.ShouldImport(item.Name ?? "", item.NsdaCategory, item.Type));
        }).OrderBy(item => item.Name, StringComparer.OrdinalIgnoreCase).ToList();
    }

    public static ImportedEvent ParseResultSet(string json)
    {
        var sets = JsonSerializer.Deserialize<List<ResultSetDto>>(json, JsonOptions)
            ?? throw new InvalidDataException("Tabroom result response was empty.");
        var set = sets.FirstOrDefault() ?? throw new InvalidDataException("Tabroom result response had no result set.");
        var eventName = set.Event?.Name ?? "Event";
        var kind = EventClassifier.Classify(eventName, set.Event?.NsdaCategory, set.Event?.Type)
            ?? throw new InvalidDataException($"Event '{eventName}' is not an MDTA debate event.");
        var headerText = HeaderText(set);
        var rows = set.Results ?? new List<ResultRowDto>();
        var seedSheet = string.Equals(set.Tag, "seed", StringComparison.OrdinalIgnoreCase);
        var breaksToOctos = headerText.Contains("octo", StringComparison.OrdinalIgnoreCase);
        var hasSemis = headerText.Contains("semi", StringComparison.OrdinalIgnoreCase);
        var hasNonPrelim = rows.Any(row => !RowIsPrelim(row));
        var hasFinalSession = hasNonPrelim || headerText.Contains("elim", StringComparison.OrdinalIgnoreCase) || headerText.Contains("final", StringComparison.OrdinalIgnoreCase);
        var schools = rows.Select(row => row.School?.Id).Where(id => id.HasValue).Select(id => id!.Value).Distinct().Count();

        var entries = rows.Select(row =>
        {
            var prelim = RowIsPrelim(row);
            var name = row.Entry?.Name;
            if (string.IsNullOrWhiteSpace(name))
            {
                name = row.Entry?.Code ?? "";
            }

            return new ImportedEntry(
                row.Entry?.Id,
                name,
                row.School?.Id,
                row.School?.Name ?? row.School?.Code ?? "Unknown school",
                row.Rank,
                prelim,
                ReachedElims: !prelim && (breaksToOctos || hasFinalSession),
                ElimWins: 0,
                ReachedFinals: !prelim,
                ReachedSemis: !prelim && hasSemis,
                NameNormalizer.SplitCompetitors(name).Select(competitor => new ImportedCompetitor(competitor)).ToList());
        }).ToList();

        return new ImportedEvent(
            set.Event?.Id ?? 0,
            set.Id,
            eventName,
            kind.ToString(),
            EventClassifier.IsCountedDivision(eventName),
            schools,
            entries.Count,
            breaksToOctos,
            hasFinalSession && kind == EventKind.Congress,
            hasSemis,
            EventClassifier.IsHouse(eventName),
            json,
            entries);

        bool RowIsPrelim(ResultRowDto row) => seedSheet || IsPrelim(row);
    }

    public static DateOnly LocalDate(DateTimeOffset start, string? timeZoneId)
    {
        TimeZoneInfo zone;
        if (!string.IsNullOrWhiteSpace(timeZoneId) && TimeZoneInfo.TryFindSystemTimeZoneById(timeZoneId, out var found))
        {
            zone = found;
        }
        else if (!TimeZoneInfo.TryFindSystemTimeZoneById("America/Chicago", out zone!))
        {
            zone = TimeZoneInfo.Utc;
        }

        return DateOnly.FromDateTime(TimeZoneInfo.ConvertTime(start, zone).DateTime);
    }

    private static int? ChooseResultSet(List<ResultSetRefDto>? sets)
    {
        if (sets is null || sets.Count == 0)
        {
            return null;
        }

        var final = sets.FirstOrDefault(set => TagIs(set, "final"));
        if (final is not null)
        {
            return final.Id;
        }

        var seed = sets.FirstOrDefault(set => TagIs(set, "seed"));
        if (seed is not null)
        {
            return seed.Id;
        }

        var other = sets.FirstOrDefault(set => !TagIs(set, "bracket") && !TagIs(set, "speaker") && !TagIs(set, "table"));
        return (other ?? sets[0]).Id;
    }

    private static bool TagIs(ResultSetRefDto set, string tag) =>
        string.Equals(set.Tag, tag, StringComparison.OrdinalIgnoreCase);

    private static bool IsPrelim(ResultRowDto row) =>
        string.Equals(row.Place, "Prelim", StringComparison.OrdinalIgnoreCase);

    private static string HeaderText(ResultSetDto set)
    {
        if (set.Headers is null)
        {
            return "";
        }

        return string.Join(' ', set.Headers.Values.Select(header =>
            $"{header.Tag} {header.Description} {header.Protocol?.Name}"));
    }

    private sealed class TournamentDto
    {
        public int Id { get; set; }
        public string? Name { get; set; }
        public string? State { get; set; }
        public string? Tz { get; set; }
        public DateTimeOffset Start { get; set; }
    }

    private sealed class EventDto
    {
        public int Id { get; set; }
        public string? Name { get; set; }
        public string? Abbr { get; set; }
        public string? Level { get; set; }
        public string? Type { get; set; }
        [JsonPropertyName("nsdacategory")]
        public int? NsdaCategory { get; set; }
        public List<ResultSetRefDto>? ResultSets { get; set; }
    }

    private sealed class ResultSetRefDto
    {
        public int Id { get; set; }
        public string? Tag { get; set; }
    }

    private sealed class ResultSetDto
    {
        public int Id { get; set; }
        public string? Tag { get; set; }
        public EventDto? Event { get; set; }
        public Dictionary<string, HeaderDto>? Headers { get; set; }
        public List<ResultRowDto>? Results { get; set; }
    }

    private sealed class HeaderDto
    {
        public string? Tag { get; set; }
        public string? Description { get; set; }
        public ProtocolDto? Protocol { get; set; }
    }

    private sealed class ProtocolDto
    {
        public string? Name { get; set; }
    }

    private sealed class ResultRowDto
    {
        public int Rank { get; set; }
        public string? Place { get; set; }
        public EntryDto? Entry { get; set; }
        public SchoolDto? School { get; set; }
    }

    private sealed class EntryDto
    {
        public int Id { get; set; }
        public string? Code { get; set; }
        public string? Name { get; set; }
    }

    private sealed class SchoolDto
    {
        public int Id { get; set; }
        public string? Code { get; set; }
        public string? Name { get; set; }
    }
}
