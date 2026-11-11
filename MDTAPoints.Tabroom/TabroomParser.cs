using System.Text.Json;
using System.Text.Json.Serialization;
using MDTAPoints.Core;

namespace MDTAPoints.Tabroom;

public static class TabroomParser
{
    private static readonly string[] ElimOrder = ["quads", "triples", "doubles", "octos", "quarters", "semis", "finals"];
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true,
        NumberHandling = JsonNumberHandling.AllowReadingFromString
    };

    public static bool IncludeSchool(string? stateCode) =>
        string.IsNullOrWhiteSpace(stateCode)
        || string.Equals(stateCode.Trim(), "MN", StringComparison.OrdinalIgnoreCase);

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
            var resultIds = ChooseResultSets(item);
            return new TabroomEventSummary(
                item.Id,
                item.Name ?? "",
                item.Abbr,
                item.Level,
                item.Type,
                item.NsdaCategory,
                resultIds.Count == 0 ? null : resultIds[^1],
                EventClassifier.ShouldImport(item.Name ?? "", item.NsdaCategory, item.Type, item.Level),
                resultIds);
        }).OrderBy(item => item.Name, StringComparer.OrdinalIgnoreCase).ToList();
    }

    public static ImportedEvent ParseResultSet(string json) => ParseResultSets(new[] { json });

    public static ImportedEvent ParseResultSets(IReadOnlyList<string> jsons)
    {
        if (jsons.Count == 0)
        {
            throw new InvalidDataException("Tabroom result response was empty.");
        }

        var sets = jsons.Select(ReadResultSet).ToList();
        if (sets.Count > 1 && sets.All(set => string.Equals(set.Tag, "chamber", StringComparison.OrdinalIgnoreCase)))
        {
            return ParseChambers(sets, string.Join('\n', jsons));
        }

        return ParseSingle(sets[0], jsons[0]);
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

    private static ImportedEvent ParseSingle(ResultSetDto set, string rawJson)
    {
        var eventName = set.Event?.Name ?? "Event";
        var kind = EventClassifier.Classify(eventName, set.Event?.NsdaCategory, set.Event?.Type)
            ?? throw new InvalidDataException($"Event '{eventName}' is not an MDTA debate event.");
        var division = EventClassifier.ClassifyDivision(eventName, set.Event?.Level);
        var headerText = HeaderText(set);
        var rows = set.Results ?? new List<ResultRowDto>();
        var seedSheet = string.Equals(set.Tag, "seed", StringComparison.OrdinalIgnoreCase);
        var roundKey = RoundHeaderKey(set);
        var rounds = rows.Select(row => RoundLabel(row, roundKey)).ToList();
        var hasElimLadder = rounds.Any(round => round is "quads" or "triples" or "doubles" or "octos" or "quarters" or "semis");
        var present = ElimOrder.Where(round => rounds.Contains(round)).ToList();
        var breaksToOctos = headerText.Contains("octo", StringComparison.OrdinalIgnoreCase)
            || present.Any(round => round is "quads" or "triples" or "doubles" or "octos");
        var hasSemis = headerText.Contains("semi", StringComparison.OrdinalIgnoreCase) || present.Contains("semis");
        var hasNonPrelim = rows.Any(row => !PrelimRow(row));
        var hasFinalSession = hasNonPrelim || headerText.Contains("elim", StringComparison.OrdinalIgnoreCase) || headerText.Contains("final", StringComparison.OrdinalIgnoreCase);
        var schools = rows.Select(row => row.School?.Id).Where(id => id.HasValue).Select(id => id!.Value).Distinct().Count();

        var entries = rows.Select(row =>
        {
            var name = EntryName(row);
            var round = RoundLabel(row, roundKey);
            var prelim = PrelimRow(row);
            int elimWins = 0;
            bool reachedElims;
            bool reachedFinals;
            bool reachedSemis;
            if (!hasElimLadder)
            {
                reachedElims = !prelim && (breaksToOctos || hasFinalSession);
                reachedFinals = !prelim;
                reachedSemis = !prelim && hasSemis;
            }
            else
            {
                prelim = seedSheet || round is null or "prelim";
                var champion = !prelim && IsChampion(row.Place);
                elimWins = prelim || round is null ? 0 : ElimWins(present, round, champion);
                reachedElims = !prelim;
                reachedFinals = round == "finals";
                reachedSemis = round is "semis" or "finals";
            }

            return new ImportedEntry(
                row.Entry?.Id,
                name,
                row.School?.Id,
                row.School?.Name ?? row.School?.Code ?? "Unknown school",
                row.Rank,
                prelim,
                reachedElims,
                elimWins,
                reachedFinals,
                reachedSemis,
                NameNormalizer.SplitCompetitors(name).Select(competitor => new ImportedCompetitor(competitor)).ToList(),
                row.School?.State);
        }).ToList();

        return new ImportedEvent(
            set.Event?.Id ?? 0,
            set.Id,
            eventName,
            kind.ToString(),
            division == Division.Varsity,
            division.ToString(),
            schools,
            entries.Count,
            breaksToOctos,
            hasFinalSession && kind == EventKind.Congress,
            hasSemis,
            EventClassifier.IsHouse(eventName),
            rawJson,
            entries);

        bool PrelimRow(ResultRowDto item) => seedSheet || IsPrelim(item);
    }

    private static ImportedEvent ParseChambers(IReadOnlyList<ResultSetDto> sets, string rawJson)
    {
        var prelim = sets.FirstOrDefault(set => ChamberKind(set) == "prelim") ?? sets[0];
        var semi = sets.FirstOrDefault(set => ChamberKind(set) == "semi");
        var final = sets.FirstOrDefault(set => ChamberKind(set) == "final");
        var source = final ?? semi ?? prelim;
        var eventName = source.Event?.Name ?? prelim.Event?.Name ?? "Congressional Debate";
        var kind = EventClassifier.Classify(eventName, source.Event?.NsdaCategory, source.Event?.Type)
            ?? throw new InvalidDataException($"Event '{eventName}' is not an MDTA debate event.");
        var division = EventClassifier.ClassifyDivision(eventName, source.Event?.Level);
        var prelimRows = IndexRows(prelim);
        var semiRows = semi is null ? new Dictionary<int, ResultRowDto>() : IndexRows(semi);
        var finalRows = final is null ? new Dictionary<int, ResultRowDto>() : IndexRows(final);
        var ids = prelimRows.Keys.Concat(semiRows.Keys).Concat(finalRows.Keys).Distinct().ToList();

        var staged = ids.Select(id =>
        {
            var inFinal = finalRows.ContainsKey(id);
            var inSemi = inFinal || semiRows.ContainsKey(id);
            var row = inFinal ? finalRows[id] : inSemi && semiRows.TryGetValue(id, out var semiRow) ? semiRow : prelimRows[id];
            var name = EntryName(row);
            var stage = inFinal ? 0 : inSemi ? 1 : 2;
            var stagePlace = inFinal
                ? NumericPlace(finalRows[id])
                : inSemi && semiRows.TryGetValue(id, out var rankedSemi)
                    ? NumericPlace(rankedSemi)
                    : prelimRows.TryGetValue(id, out var rankedPrelim) ? NumericPlace(rankedPrelim) : row.Rank;
            return new StagedEntry(id, name, row, stage, stagePlace, inSemi, inFinal);
        }).ToList();

        var maxFinal = staged.Where(item => item.Stage == 0).Select(item => item.StagePlace).DefaultIfEmpty(0).Max();
        var place = maxFinal;
        var entries = staged
            .OrderBy(item => item.Stage)
            .ThenBy(item => item.StagePlace)
            .ThenBy(item => item.Name, StringComparer.OrdinalIgnoreCase)
            .Select(item =>
            {
                var assigned = item.Stage == 0 ? item.StagePlace : ++place;
                return new ImportedEntry(
                    item.Id,
                    item.Name,
                    item.Row.School?.Id,
                    item.Row.School?.Name ?? item.Row.School?.Code ?? "Unknown school",
                    assigned,
                    !item.InSemi && !item.InFinal,
                    item.InSemi || item.InFinal,
                    0,
                    item.InFinal,
                    item.InSemi,
                    NameNormalizer.SplitCompetitors(item.Name).Select(competitor => new ImportedCompetitor(competitor)).ToList(),
                    item.Row.School?.State);
            })
            .ToList();

        var schools = entries.Select(entry => entry.TabroomSchoolId).Where(id => id.HasValue).Select(id => id!.Value).Distinct().Count();
        return new ImportedEvent(
            source.Event?.Id ?? prelim.Event?.Id ?? 0,
            (final ?? semi ?? prelim).Id,
            eventName,
            kind.ToString(),
            division == Division.Varsity,
            division.ToString(),
            schools,
            entries.Count,
            false,
            finalRows.Count > 0 && kind == EventKind.Congress,
            semiRows.Count > 0,
            EventClassifier.IsHouse(eventName),
            rawJson,
            entries);
    }

    private static List<int> ChooseResultSets(EventDto item)
    {
        var sets = item.ResultSets;
        if (sets is null || sets.Count == 0)
        {
            return new List<int>();
        }

        var final = sets.FirstOrDefault(set => TagIs(set, "final"));
        if (final is not null)
        {
            return new List<int> { final.Id };
        }

        var seed = sets.FirstOrDefault(set => TagIs(set, "seed"));
        if (seed is not null)
        {
            return new List<int> { seed.Id };
        }

        var chambers = sets.Where(set => TagIs(set, "chamber")).ToList();
        if (chambers.Count > 0 && EventClassifier.Classify(item.Name ?? "", item.NsdaCategory, item.Type) == EventKind.Congress)
        {
            return chambers
                .OrderBy(set => ChamberRank(set.Label))
                .Select(set => set.Id)
                .ToList();
        }

        var other = sets.FirstOrDefault(set => !TagIs(set, "bracket") && !TagIs(set, "speaker") && !TagIs(set, "table"));
        return new List<int> { (other ?? sets[0]).Id };
    }

    private static int ChamberRank(string? label)
    {
        var text = label ?? "";
        if (text.Contains("semi", StringComparison.OrdinalIgnoreCase))
        {
            return 1;
        }

        if (text.Contains("final", StringComparison.OrdinalIgnoreCase))
        {
            return 2;
        }

        return 0;
    }

    private static string ChamberKind(ResultSetDto set)
    {
        var label = set.Label ?? "";
        if (label.Contains("semi", StringComparison.OrdinalIgnoreCase))
        {
            return "semi";
        }

        if (label.Contains("final", StringComparison.OrdinalIgnoreCase))
        {
            return "final";
        }

        return "prelim";
    }

    private static Dictionary<int, ResultRowDto> IndexRows(ResultSetDto set)
    {
        var rows = new Dictionary<int, ResultRowDto>();
        var missingId = -1;
        foreach (var row in set.Results ?? new List<ResultRowDto>())
        {
            var id = row.Entry?.Id ?? 0;
            if (id == 0)
            {
                id = missingId--;
            }

            if (rows.ContainsKey(id))
            {
                continue;
            }

            rows[id] = row;
        }

        return rows;
    }

    private static int NumericPlace(ResultRowDto row) =>
        int.TryParse(row.Place, out var place) ? place : row.Rank;

    private static string EntryName(ResultRowDto row)
    {
        var name = row.Entry?.Name;
        return string.IsNullOrWhiteSpace(name) ? row.Entry?.Code ?? "" : name;
    }

    private static int ElimWins(List<string> present, string round, bool champion)
    {
        var index = present.IndexOf(round);
        if (index < 0)
        {
            return 0;
        }

        return champion ? index + 1 : index;
    }

    private static bool IsChampion(string? place) =>
        string.Equals(place?.Trim(), "1st", StringComparison.OrdinalIgnoreCase)
        || string.Equals(place?.Trim(), "first", StringComparison.OrdinalIgnoreCase);

    private static string? RoundHeaderKey(ResultSetDto set)
    {
        if (set.Headers is null)
        {
            return null;
        }

        foreach (var header in set.Headers)
        {
            if (string.Equals(header.Value.Tag, "Round", StringComparison.OrdinalIgnoreCase)
                || (header.Value.Description?.Contains("last round", StringComparison.OrdinalIgnoreCase) ?? false))
            {
                return header.Key;
            }
        }

        return null;
    }

    private static string? RoundLabel(ResultRowDto row, string? roundKey)
    {
        var fromPlace = ClassifyRound(row.Place);
        if (fromPlace is not null)
        {
            return fromPlace;
        }

        if (roundKey is null || row.Values is null || !row.Values.TryGetValue(roundKey, out var value))
        {
            return null;
        }

        return ClassifyRound(ValueText(value));
    }

    private static string? ValueText(JsonElement element) =>
        element.ValueKind switch
        {
            JsonValueKind.String => element.GetString(),
            JsonValueKind.Number => element.GetRawText(),
            _ => null
        };

    private static string? ClassifyRound(string? place)
    {
        if (string.IsNullOrWhiteSpace(place))
        {
            return null;
        }

        var text = place.Trim().ToLowerInvariant();
        if (text is "prelim" or "prelims")
        {
            return "prelim";
        }

        if (text.Contains("quad"))
        {
            return "quads";
        }

        if (text.Contains("triple"))
        {
            return "triples";
        }

        if (text.Contains("double"))
        {
            return "doubles";
        }

        if (text.Contains("octo"))
        {
            return "octos";
        }

        if (text.Contains("quarter"))
        {
            return "quarters";
        }

        if (text.Contains("semi"))
        {
            return "semis";
        }

        if (text is "1st" or "2nd" or "first" or "second" || text.Contains("final"))
        {
            return "finals";
        }

        return null;
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

    private static ResultSetDto ReadResultSet(string json)
    {
        var sets = JsonSerializer.Deserialize<List<ResultSetDto>>(json, JsonOptions)
            ?? throw new InvalidDataException("Tabroom result response was empty.");
        return sets.FirstOrDefault() ?? throw new InvalidDataException("Tabroom result response had no result set.");
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
        public string? Label { get; set; }
    }

    private sealed class ResultSetDto
    {
        public int Id { get; set; }
        public string? Tag { get; set; }
        public string? Label { get; set; }
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
        public Dictionary<string, JsonElement>? Values { get; set; }
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
        public string? State { get; set; }
    }

    private sealed record StagedEntry(
        int Id,
        string Name,
        ResultRowDto Row,
        int Stage,
        int StagePlace,
        bool InSemi,
        bool InFinal);
}
