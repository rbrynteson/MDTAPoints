using MDTAPoints.Api.Services;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;

namespace MDTAPoints.Api.Pages.Admin.Tournaments;

public class ImportModel : PageModel
{
    private readonly TabroomImportService _imports;

    public ImportModel(TabroomImportService imports)
    {
        _imports = imports;
    }

    [BindProperty]
    public int TournamentId { get; set; }

    public string? Error { get; private set; }

    public void OnGet()
    {
    }

    public async Task<IActionResult> OnPostAsync()
    {
        try
        {
            var outcome = await _imports.ImportAsync(TournamentId, HttpContext.RequestAborted);
            if (outcome.AlreadyPublished)
            {
                return RedirectToPage("Index", new { message = "That tournament is already published." });
            }

            return RedirectToPage("Review", new { id = outcome.TournamentId });
        }
        catch (Exception ex) when (ex is HttpRequestException or InvalidDataException)
        {
            Error = "Tabroom did not return results for that tournament id.";
            return Page();
        }
    }
}
