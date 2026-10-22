using MDTAPoints.Api.Data;
using MDTAPoints.Api.Services;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;

namespace MDTAPoints.Api.Pages.Admin.Tournaments;

public class IndexModel : PageModel
{
    private readonly CupDatabase _database;

    public IndexModel(CupDatabase database)
    {
        _database = database;
    }

    public IActionResult OnGet(string? message) =>
        RedirectToPage("/Tournaments/Index", new { message });

    public async Task<IActionResult> OnGetExportAsync()
    {
        var tournaments = await _database.ListTournamentsAsync(HttpContext.RequestAborted);
        return File(ExcelExports.Tournaments(tournaments), ExcelExports.ContentType, "mdta-tournaments.xlsx");
    }
}
