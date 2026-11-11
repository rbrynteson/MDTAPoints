using MDTAPoints.Api.Data;
using Microsoft.AspNetCore.Mvc.RazorPages;

namespace MDTAPoints.Api.Pages.Schools;

public class IndexModel : PageModel
{
    private readonly CupDatabase _database;

    public IndexModel(CupDatabase database)
    {
        _database = database;
    }

    public IReadOnlyList<SchoolListItem> Schools { get; private set; } = Array.Empty<SchoolListItem>();

    public async Task OnGetAsync()
    {
        Schools = await _database.ListSchoolsAsync(HttpContext.RequestAborted);
    }
}
