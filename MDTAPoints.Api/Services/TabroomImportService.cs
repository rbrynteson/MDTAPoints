using MDTAPoints.Api.Data;
using MDTAPoints.Tabroom;

namespace MDTAPoints.Api.Services;

public sealed class TabroomImportService
{
    private readonly TabroomClient _tabroom;
    private readonly CupDatabase _database;

    public TabroomImportService(TabroomClient tabroom, CupDatabase database)
    {
        _tabroom = tabroom;
        _database = database;
    }

    public async Task<ImportOutcome> ImportAsync(int tabroomTournamentId, CancellationToken cancellationToken)
    {
        var existing = await _database.FindByTabroomIdAsync(tabroomTournamentId, cancellationToken);
        if (existing is { Status: "Published" })
        {
            return new ImportOutcome(existing.Id, true);
        }

        var tournamentJson = await _tabroom.GetTournamentJsonAsync(tabroomTournamentId, cancellationToken);
        var info = TabroomParser.ParseTournament(tournamentJson);
        var summaries = TabroomParser.ParseEvents(await _tabroom.GetEventsJsonAsync(tabroomTournamentId, cancellationToken));
        var imported = new List<ImportedEvent>();
        var skipped = new List<string>();
        foreach (var summary in summaries)
        {
            if (!summary.ShouldImport)
            {
                skipped.Add(summary.Name);
                continue;
            }

            if (summary.ResultSetId is null)
            {
                skipped.Add($"{summary.Name} (no results posted)");
                continue;
            }

            var resultJson = await _tabroom.GetResultJsonAsync(tabroomTournamentId, summary.ResultSetId.Value, cancellationToken);
            imported.Add(TabroomParser.ParseResultSet(resultJson));
        }

        var id = await _database.SaveTabroomImportAsync(
            info,
            imported,
            string.Join(Environment.NewLine, skipped),
            existing?.Id,
            cancellationToken);
        return new ImportOutcome(id, false);
    }
}
