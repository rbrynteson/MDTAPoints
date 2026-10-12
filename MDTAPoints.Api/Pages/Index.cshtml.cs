using MDTAPoints.Api.Data;
using MDTAPoints.Core;
using Microsoft.AspNetCore.Mvc.RazorPages;

namespace MDTAPoints.Api.Pages;

public class IndexModel : PageModel
{
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
}
