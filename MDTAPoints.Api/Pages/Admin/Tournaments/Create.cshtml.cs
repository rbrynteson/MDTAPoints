using MDTAPoints.Api.Data;
using MDTAPoints.Core;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;

namespace MDTAPoints.Api.Pages.Admin.Tournaments;

public class CreateModel : PageModel
{
    private readonly CupDatabase _database;

    public CreateModel(CupDatabase database)
    {
        _database = database;
    }

    [BindProperty]
    public string Name { get; set; } = "";

    [BindProperty]
    public DateOnly StartDate { get; set; } = DateOnly.FromDateTime(DateTime.Today);

    [BindProperty]
    public string State { get; set; } = "MN";

    [BindProperty]
    public string Kind { get; set; } = TournamentKind.Invitational.ToString();

    public async Task<IActionResult> OnPostAsync()
    {
        var id = await _database.CreateManualAsync(Name, StartDate, State, Kind, HttpContext.RequestAborted);
        return RedirectToPage("Review", new { id });
    }
}
