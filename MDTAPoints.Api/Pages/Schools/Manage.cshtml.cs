using MDTAPoints.Api.Data;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;

namespace MDTAPoints.Api.Pages.Schools;

[Authorize]
public class ManageModel : PageModel
{
    private readonly CupDatabase _database;

    public ManageModel(CupDatabase database)
    {
        _database = database;
    }

    public SchoolListItem? School { get; private set; }

    public async Task<IActionResult> OnGetAsync(int id)
    {
        School = await _database.GetSchoolAsync(id, HttpContext.RequestAborted);
        return School is null ? NotFound() : Page();
    }

    public async Task<IActionResult> OnPostAsync(int id, bool exclude)
    {
        if (!await _database.SetSchoolExcludedAsync(id, exclude, HttpContext.RequestAborted))
        {
            return NotFound();
        }

        return RedirectToPage(new { id });
    }
}
