using MDTAPoints.Api.Data;
using MDTAPoints.Api.Services;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;

namespace MDTAPoints.Api.Pages.Tournaments;

public class ResultsModel : PageModel
{
    private readonly CupDatabase _database;

    public ResultsModel(CupDatabase database)
    {
        _database = database;
    }

    public TournamentDetail? Tournament { get; private set; }
    public IReadOnlyList<EventRow> Events { get; private set; } = Array.Empty<EventRow>();
    public EventRow? Selected { get; private set; }
    public IReadOnlyList<string> Rules { get; private set; } = Array.Empty<string>();
    public IReadOnlyList<TournamentResultRow> Rows { get; private set; } = Array.Empty<TournamentResultRow>();

    public async Task<IActionResult> OnGetAsync(int id, int? eventId)
    {
        if (!await LoadAsync(id, eventId))
        {
            return NotFound();
        }

        return Page();
    }

    public async Task<IActionResult> OnGetExportAsync(int id, int? eventId)
    {
        if (!await LoadAsync(id, eventId) || Tournament is null || Selected is null)
        {
            return NotFound();
        }

        var title = TournamentResults.EventTitle(Selected);
        var fileName = $"mdta-{ExcelExports.FileSlug(Tournament.Name)}-{ExcelExports.FileSlug(title)}.xlsx";
        return File(ExcelExports.EventResults(title, Rows), ExcelExports.ContentType, fileName);
    }

    private async Task<bool> LoadAsync(int id, int? eventId)
    {
        Tournament = await _database.GetTournamentAsync(id, HttpContext.RequestAborted);
        if (Tournament is null)
        {
            return false;
        }

        Events = TournamentResults.Order(Tournament.Events);
        Selected = Events.FirstOrDefault(item => item.Id == eventId) ?? Events.FirstOrDefault();
        if (Selected is null)
        {
            Rules = Array.Empty<string>();
            Rows = Array.Empty<TournamentResultRow>();
            return true;
        }

        Rules = TournamentResults.Rules(Tournament, Selected);
        Rows = TournamentResults.Rows(Tournament, Selected);
        return true;
    }
}
