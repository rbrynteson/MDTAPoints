using MDTAPoints.Api.Data;
using MDTAPoints.Api.Services;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;

namespace MDTAPoints.Api.Pages.Admin.Tournaments;

public class ReviewModel : PageModel
{
    private readonly CupDatabase _database;

    public ReviewModel(CupDatabase database)
    {
        _database = database;
    }

    public TournamentDetail? Tournament { get; private set; }
    public ILookup<int, SchoolStudent> Students { get; private set; } = Array.Empty<SchoolStudent>().ToLookup(student => student.SchoolId);
    public string? Error { get; private set; }
    private string? PlaceError { get; set; }

    public IEnumerable<SchoolStudent> StudentsAt(int schoolId) => Students[schoolId];

    public async Task<IActionResult> OnGetAsync(int id)
    {
        await LoadAsync(id);
        return Tournament is null ? NotFound() : Page();
    }

    public async Task<IActionResult> OnGetExportAsync(int id)
    {
        var tournament = await _database.GetTournamentAsync(id, HttpContext.RequestAborted);
        if (tournament is null)
        {
            return NotFound();
        }

        var fileName = $"mdta-{ExcelExports.FileSlug(tournament.Name)}.xlsx";
        return File(ExcelExports.Tournament(tournament), ExcelExports.ContentType, fileName);
    }

    public async Task<IActionResult> OnPostSaveAsync(int id)
    {
        var update = ReadUpdate();
        if (PlaceError is not null)
        {
            await LoadAsync(id);
            Error = PlaceError;
            return Page();
        }

        await _database.SaveReviewAsync(id, update, HttpContext.RequestAborted);
        return RedirectToPage(new { id });
    }

    public async Task<IActionResult> OnPostPublishAsync(int id)
    {
        var update = ReadUpdate();
        if (PlaceError is not null)
        {
            await LoadAsync(id);
            Error = PlaceError;
            return Page();
        }

        await _database.SaveReviewAsync(id, update, HttpContext.RequestAborted);
        try
        {
            await _database.PublishAsync(id, HttpContext.RequestAborted);
        }
        catch (InvalidOperationException ex)
        {
            await LoadAsync(id);
            Error = ex.Message;
            return Page();
        }

        return RedirectToPage("Index", new { message = "Tournament published." });
    }

    public async Task<IActionResult> OnPostAddAsync(int id, string eventKind, string eventName, int place, string school, string competitors, int elimWins)
    {
        await _database.AddEntryAsync(
            id,
            eventKind,
            eventName,
            place,
            school,
            competitors,
            Request.Form.ContainsKey("prelim"),
            elimWins,
            Request.Form.ContainsKey("reachedFinals"),
            false,
            false,
            HttpContext.RequestAborted);
        return RedirectToPage(new { id });
    }

    private async Task LoadAsync(int id)
    {
        Tournament = await _database.GetTournamentAsync(id, HttpContext.RequestAborted);
        if (Tournament is null)
        {
            return;
        }

        var schoolIds = Tournament.Events.SelectMany(item => item.Entries).Select(item => item.SchoolId);
        Students = (await _database.StudentsForSchoolsAsync(schoolIds, HttpContext.RequestAborted)).ToLookup(student => student.SchoolId);
    }

    private ReviewUpdate ReadUpdate()
    {
        var update = new ReviewUpdate { Kind = Request.Form["kind"].ToString() };
        var octos = Ids("octos");
        var finals = Ids("finals");
        var semis = Ids("semis");
        foreach (var eventId in Ids("eventId"))
        {
            update.Octos[eventId] = octos.Contains(eventId);
            update.FinalSessions[eventId] = finals.Contains(eventId);
            update.Semis[eventId] = semis.Contains(eventId);
        }

        foreach (var key in Request.Form.Keys)
        {
            if (key.StartsWith("elim_", StringComparison.Ordinal) && int.TryParse(key["elim_".Length..], out var entryId))
            {
                var entry = update.Entries.TryGetValue(entryId, out var existing) ? existing : new EntryEdit();
                entry.ElimWins = int.TryParse(Request.Form[key], out var wins) ? wins : 0;
                entry.ReachedElims = Request.Form.ContainsKey($"reached_{entryId}");
                entry.ReachedFinals = Request.Form.ContainsKey($"finalist_{entryId}");
                entry.ReachedSemis = Request.Form.ContainsKey($"semi_{entryId}");
                entry.QualifiedToState = Request.Form.ContainsKey($"qualified_{entryId}");
                var rawPlace = Request.Form[$"place_{entryId}"].ToString();
                if (int.TryParse(rawPlace, out var place) && place >= 1)
                {
                    entry.Place = place;
                }
                else if (!string.IsNullOrWhiteSpace(rawPlace))
                {
                    PlaceError = "Each place must be a whole number of 1 or higher.";
                }

                var note = Request.Form[$"note_{entryId}"].ToString().Trim();
                entry.PlaceNote = note.Length == 0 ? null : note.Length <= 400 ? note : note[..400];
                update.Entries[entryId] = entry;
            }

            if (key.StartsWith("link_", StringComparison.Ordinal) && int.TryParse(key["link_".Length..], out var competitorId))
            {
                update.CompetitorChoices[competitorId] = Request.Form[key].ToString();
            }
        }

        return update;
    }

    private HashSet<int> Ids(string key) =>
        Request.Form[key]
            .Select(value => int.TryParse(value, out var id) ? id : (int?)null)
            .Where(id => id.HasValue)
            .Select(id => id!.Value)
            .ToHashSet();
}
