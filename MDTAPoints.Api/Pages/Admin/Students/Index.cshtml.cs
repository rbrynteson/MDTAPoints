using MDTAPoints.Api.Data;
using MDTAPoints.Api.Services;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;

namespace MDTAPoints.Api.Pages.Admin.Students;

public class IndexModel : PageModel
{
    private readonly CupDatabase _database;

    public IndexModel(CupDatabase database)
    {
        _database = database;
    }

    public IReadOnlyList<StudentAdminRow> Students { get; private set; } = Array.Empty<StudentAdminRow>();
    public IReadOnlyList<HeadToHeadRow> HeadToHead { get; private set; } = Array.Empty<HeadToHeadRow>();
    public string? Message { get; private set; }

    public async Task OnGetAsync(string? message)
    {
        Message = message;
        await LoadAsync();
    }

    public async Task<IActionResult> OnPostSaveAsync(int id, int? grade)
    {
        var events = Request.Form["qualified"].ToArray().Select(value => value ?? "").Where(value => value.Length > 0).ToArray();
        await _database.UpdateStudentAsync(id, grade, events, HttpContext.RequestAborted);
        return RedirectToPage(new { message = "Student saved." });
    }

    public async Task<IActionResult> OnPostHeadToHeadAsync(int winnerId, int loserId)
    {
        await _database.AddHeadToHeadAsync(winnerId, loserId, HttpContext.RequestAborted);
        return RedirectToPage();
    }

    public async Task<IActionResult> OnPostDeleteHeadToHeadAsync(int winnerId, int loserId)
    {
        await _database.DeleteHeadToHeadAsync(winnerId, loserId, HttpContext.RequestAborted);
        return RedirectToPage();
    }

    public async Task<IActionResult> OnGetExportAsync()
    {
        var students = await _database.ListStudentsAsync(HttpContext.RequestAborted);
        var headToHead = await _database.ListHeadToHeadAsync(HttpContext.RequestAborted);
        return File(ExcelExports.Students(students, headToHead), ExcelExports.ContentType, "mdta-students.xlsx");
    }

    private async Task LoadAsync()
    {
        Students = await _database.ListStudentsAsync(HttpContext.RequestAborted);
        HeadToHead = await _database.ListHeadToHeadAsync(HttpContext.RequestAborted);
    }
}
