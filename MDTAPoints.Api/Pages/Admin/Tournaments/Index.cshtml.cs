using MDTAPoints.Api.Data;
using Microsoft.AspNetCore.Mvc.RazorPages;

namespace MDTAPoints.Api.Pages.Admin.Tournaments;

public class IndexModel : PageModel
{
    private readonly CupDatabase _database;

    public IndexModel(CupDatabase database)
    {
        _database = database;
    }

    public IReadOnlyList<TournamentListItem> Tournaments { get; private set; } = Array.Empty<TournamentListItem>();
    public string? Message { get; private set; }

    public async Task OnGetAsync(string? message)
    {
        Message = message;
        Tournaments = await _database.ListTournamentsAsync(HttpContext.RequestAborted);
    }
}
