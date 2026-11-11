using MDTAPoints.Api.Data;
using MDTAPoints.Api.Services;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;

namespace MDTAPoints.Api.Pages.Tournaments;

public class IndexModel : PageModel
{
    private readonly CupDatabase _database;

    public IndexModel(CupDatabase database)
    {
        _database = database;
    }

    public IReadOnlyList<TournamentListItem> Tournaments { get; private set; } = Array.Empty<TournamentListItem>();
    public IReadOnlyDictionary<int, int> NamesToReview { get; private set; } = new Dictionary<int, int>();
    public string? Message { get; private set; }

    public int NamesWaiting(int tournamentId) =>
        NamesToReview.TryGetValue(tournamentId, out var count) ? count : 0;

    public async Task OnGetAsync(string? message)
    {
        Message = message;
        Tournaments = await _database.ListTournamentsAsync(HttpContext.RequestAborted);
        if (User.Identity?.IsAuthenticated == true)
        {
            NamesToReview = await _database.PossibleNameCountsAsync(HttpContext.RequestAborted);
        }
    }

    public async Task<IActionResult> OnGetExportAsync()
    {
        var tournaments = await _database.ListTournamentsAsync(HttpContext.RequestAborted);
        return File(ExcelExports.Tournaments(tournaments), ExcelExports.ContentType, "mdta-tournaments.xlsx");
    }
}
