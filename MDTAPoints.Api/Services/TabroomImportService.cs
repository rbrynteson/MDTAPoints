using MDTAPoints.Api.Data;
using MDTAPoints.Core;
using MDTAPoints.Tabroom;

namespace MDTAPoints.Api.Services;

public sealed record TabroomImportPreview(
    TabroomTournamentInfo? Tournament,
    IReadOnlyList<TabroomEventSummary> Events,
    TournamentListItem? Existing);

public sealed class TabroomImportService
{
    private readonly TabroomClient _tabroom;
    private readonly CupDatabase _database;

    public TabroomImportService(TabroomClient tabroom, CupDatabase database)
    {
        _tabroom = tabroom;
        _database = database;
    }

    public async Task<TabroomImportPreview> LookupAsync(int tabroomTournamentId, CancellationToken cancellationToken)
    {
        var existing = await _database.FindByTabroomIdAsync(tabroomTournamentId, cancellationToken);
        if (existing is { Status: "Published" })
        {
            return new TabroomImportPreview(null, Array.Empty<TabroomEventSummary>(), existing);
        }

        var tournamentJson = await _tabroom.GetTournamentJsonAsync(tabroomTournamentId, cancellationToken);
        var info = TabroomParser.ParseTournament(tournamentJson);
        var summaries = TabroomParser.ParseEvents(await _tabroom.GetEventsJsonAsync(tabroomTournamentId, cancellationToken));
        return new TabroomImportPreview(info, summaries, existing);
    }

    public Task<ImportOutcome> ImportAsync(int tabroomTournamentId, CancellationToken cancellationToken) =>
        ImportAsync(tabroomTournamentId, null, null, cancellationToken);

    public Task<ImportOutcome> ImportAsync(
        int tabroomTournamentId,
        IReadOnlyCollection<int>? eventIds,
        CancellationToken cancellationToken) =>
        ImportAsync(tabroomTournamentId, eventIds, null, cancellationToken);

    public async Task<ImportOutcome> ImportAsync(
        int tabroomTournamentId,
        IReadOnlyCollection<int>? eventIds,
        IReadOnlyDictionary<int, bool>? varsityByEventId,
        CancellationToken cancellationToken)
    {
        var existing = await _database.FindByTabroomIdAsync(tabroomTournamentId, cancellationToken);
        if (existing is { Status: "Published" })
        {
            return new ImportOutcome(existing.Id, true);
        }

        var tournamentJson = await _tabroom.GetTournamentJsonAsync(tabroomTournamentId, cancellationToken);
        var info = TabroomParser.ParseTournament(tournamentJson);
        var summaries = TabroomParser.ParseEvents(await _tabroom.GetEventsJsonAsync(tabroomTournamentId, cancellationToken));
        var selected = eventIds?.ToHashSet();
        var imported = new List<ImportedEvent>();
        var skipped = new List<string>();
        foreach (var summary in summaries)
        {
            var chosen = selected is null ? summary.ShouldImport : selected.Contains(summary.EventId);
            if (!chosen)
            {
                skipped.Add(summary.Name);
                continue;
            }

            if (EventClassifier.Classify(summary.Name, summary.NsdaCategory, summary.Type) is null)
            {
                skipped.Add($"{summary.Name} (not a Cup debate event)");
                continue;
            }

            var setIds = summary.ResultSetIds.Count > 0
                ? summary.ResultSetIds
                : summary.ResultSetId is int single ? new[] { single } : Array.Empty<int>();
            if (setIds.Count == 0)
            {
                skipped.Add($"{summary.Name} (no results posted)");
                continue;
            }

            var resultJson = new List<string>();
            foreach (var setId in setIds)
            {
                resultJson.Add(await _tabroom.GetResultJsonAsync(tabroomTournamentId, setId, cancellationToken));
            }

            var parsed = TabroomParser.ParseResultSets(resultJson);
            if (varsityByEventId is not null && varsityByEventId.TryGetValue(summary.EventId, out var varsity))
            {
                parsed = ApplyVarsityChoice(parsed, varsity);
            }

            parsed = OmitOutOfStateSchools(parsed, skipped, summary.Name);
            imported.Add(parsed);
        }

        if (selected is not null && imported.Count == 0)
        {
            throw new InvalidOperationException("None of the selected events have results that can be imported.");
        }

        var id = await _database.SaveTabroomImportAsync(
            info,
            imported,
            string.Join(Environment.NewLine, skipped),
            existing?.Id,
            cancellationToken);
        return new ImportOutcome(id, false);
    }

    public static ImportedEvent ApplyVarsityChoice(ImportedEvent imported, bool varsity) =>
        imported with
        {
            CountedDivision = varsity,
            Division = varsity ? Division.Varsity.ToString() : imported.Division
        };

    public static ImportedEvent OmitOutOfStateSchools(ImportedEvent imported, ICollection<string>? skipped = null, string? eventName = null)
    {
        var excluded = imported.Entries
            .Where(entry => !TabroomParser.IncludeSchool(entry.SchoolState))
            .Select(entry => string.IsNullOrWhiteSpace(entry.SchoolState) ? entry.SchoolName : $"{entry.SchoolName} ({entry.SchoolState.Trim()})")
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
            .ToList();
        if (excluded.Count > 0 && skipped is not null)
        {
            skipped.Add($"{eventName ?? imported.Name}: out-of-state schools not imported: {string.Join(", ", excluded)}");
        }

        // School and entry counts stay on the full posted field. Only Minnesota schools are saved.
        return imported with
        {
            Entries = imported.Entries.Where(entry => TabroomParser.IncludeSchool(entry.SchoolState)).ToList()
        };
    }
}
