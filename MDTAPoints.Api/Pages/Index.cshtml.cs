using MDTAPoints.Api.Data;
using MDTAPoints.Api.Services;
using MDTAPoints.Core;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;

namespace MDTAPoints.Api.Pages;

public class IndexModel : PageModel
{
    private static readonly EventKind[] EventOrder =
    {
        EventKind.LincolnDouglas,
        EventKind.Classic,
        EventKind.PublicForum,
        EventKind.Policy,
        EventKind.Congress
    };

    private readonly CupDatabase _database;

    public IndexModel(CupDatabase database)
    {
        _database = database;
    }

    public EventKind Event { get; private set; } = EventKind.LincolnDouglas;
    public IReadOnlyList<IndividualStanding> Individuals { get; private set; } = Array.Empty<IndividualStanding>();
    public IReadOnlyList<TeamStanding> Teams { get; private set; } = Array.Empty<TeamStanding>();

    public async Task OnGetAsync(string? @event)
    {
        if (Enum.TryParse<EventKind>(@event, true, out var kind))
        {
            Event = kind;
        }

        var standings = StandingsCalculator.Calculate(await _database.LoadSeasonAsync(HttpContext.RequestAborted));
        Individuals = standings.Individuals.Where(row => row.Event == Event).ToList();
        Teams = standings.Teams.Where(row => row.Event == Event).OrderBy(row => row.Rank).ToList();
    }

    public async Task<IActionResult> OnGetExportAsync(string? @event, string? table)
    {
        if (Enum.TryParse<EventKind>(@event, true, out var kind))
        {
            Event = kind;
        }

        var choice = table?.Trim().ToLowerInvariant();
        var all = choice == "all";
        if (choice is not ("individuals" or "teams" or "both" or "all"))
        {
            choice = "both";
            all = false;
        }

        var standings = StandingsCalculator.Calculate(await _database.LoadSeasonAsync(HttpContext.RequestAborted));
        var individuals = all
            ? standings.Individuals.OrderBy(row => EventIndex(row.Event)).ToList()
            : standings.Individuals.Where(row => row.Event == Event).ToList();
        var teams = all
            ? standings.Teams.OrderBy(row => EventIndex(row.Event)).ThenBy(row => row.Rank).ToList()
            : standings.Teams.Where(row => row.Event == Event).OrderBy(row => row.Rank).ToList();
        var bytes = ExcelExports.Standings(
            individuals,
            teams,
            includeIndividuals: choice != "teams",
            includeTeams: choice != "individuals",
            includeEvent: all);
        var slug = ExcelExports.FileSlug(EventKindLabels.Display(Event));
        var fileName = choice switch
        {
            "all" => "mdta-cup-standings.xlsx",
            "individuals" => $"mdta-cup-{slug}-individuals.xlsx",
            "teams" => $"mdta-cup-{slug}-teams.xlsx",
            _ => $"mdta-cup-{slug}.xlsx"
        };
        return File(bytes, ExcelExports.ContentType, fileName);
    }

    private static int EventIndex(EventKind kind)
    {
        var index = Array.IndexOf(EventOrder, kind);
        return index < 0 ? int.MaxValue : index;
    }
}
