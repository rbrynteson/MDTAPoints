using MDTAPoints.Api.Services;
using MDTAPoints.Core;
using MDTAPoints.Tabroom;
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

    [BindProperty]
    public List<int> EventIds { get; set; } = new();

    // Name must be set. Without it, a post that has no Varsity fields (Show events) falls back to an
    // empty prefix and tries to parse every field name, including TournamentId, as a dictionary key.
    [BindProperty(Name = "Varsity")]
    public Dictionary<int, bool> Varsity { get; set; } = new();

    public string? Error { get; private set; }
    public TabroomImportPreview? Preview { get; private set; }
    public IReadOnlyList<ImportEventRow> Events { get; private set; } = Array.Empty<ImportEventRow>();
    private bool _usePostedSelection;

    public bool IsChecked(ImportEventRow row) =>
        _usePostedSelection ? EventIds.Contains(row.EventId) : row.Suggested;

    public bool IsVarsity(ImportEventRow row) =>
        _usePostedSelection && Varsity.TryGetValue(row.EventId, out var posted) ? posted : row.Counts;

    public void OnGet()
    {
    }

    public async Task<IActionResult> OnPostFindAsync()
    {
        if (TournamentId < 1)
        {
            Error = "Enter a Tabroom tournament id.";
            return Page();
        }

        try
        {
            await LoadPreviewAsync();
        }
        catch (Exception ex) when (ex is HttpRequestException or InvalidDataException)
        {
            Error = "Tabroom did not return that tournament. Use the tournament id, such as 40756, not the result-set number at the end of a results link.";
            return Page();
        }

        return Page();
    }

    public async Task<IActionResult> OnPostImportAsync()
    {
        if (TournamentId < 1)
        {
            Error = "Enter a Tabroom tournament id.";
            return Page();
        }

        if (EventIds.Count == 0)
        {
            Error = "Choose at least one event.";
            _usePostedSelection = true;
            return await ReloadPreviewAsync();
        }

        try
        {
            var outcome = await _imports.ImportAsync(TournamentId, EventIds, Varsity, HttpContext.RequestAborted);
            if (outcome.AlreadyPublished)
            {
                return RedirectToPage("Index", new { message = "That tournament is already published." });
            }

            return RedirectToPage("Review", new { id = outcome.TournamentId });
        }
        catch (InvalidOperationException ex)
        {
            Error = ex.Message;
            _usePostedSelection = true;
            return await ReloadPreviewAsync();
        }
        catch (Exception ex) when (ex is HttpRequestException or InvalidDataException)
        {
            Error = "Tabroom did not return results for the selected events.";
            _usePostedSelection = true;
            return await ReloadPreviewAsync();
        }
    }

    private async Task<IActionResult> ReloadPreviewAsync()
    {
        try
        {
            await LoadPreviewAsync();
        }
        catch (Exception ex) when (ex is HttpRequestException or InvalidDataException)
        {
            Error = "Tabroom did not return that tournament. Use the tournament id, such as 40756, not the result-set number at the end of a results link.";
        }

        return Page();
    }

    private async Task LoadPreviewAsync()
    {
        Preview = await _imports.LookupAsync(TournamentId, HttpContext.RequestAborted);
        Events = Preview.Tournament is null ? Array.Empty<ImportEventRow>() : BuildRows(Preview.Events);
    }

    public static IReadOnlyList<ImportEventRow> BuildRows(IReadOnlyList<TabroomEventSummary> events) =>
        events.Select(item =>
        {
            var kind = EventClassifier.Classify(item.Name, item.NsdaCategory, item.Type);
            var hasResults = item.ResultSetId.HasValue;
            var counts = kind is not null && EventClassifier.IsCountedDivision(item.Name, item.Level);
            var suggested = kind is not null && hasResults;
            var cup = kind is null
                ? "Not a Cup event"
                : !hasResults
                    ? "No results posted"
                    : counts ? "Varsity" : "Novice or JV, not in the Cup";
            return new ImportEventRow(
                item.EventId,
                item.Name,
                kind is null ? "—" : EventKindLabels.Display(kind.Value),
                cup,
                kind is not null && hasResults,
                suggested,
                counts);
        })
        .OrderBy(row => row.Suggested ? 0 : row.Selectable ? 1 : 2)
        .ThenBy(row => row.Name, StringComparer.OrdinalIgnoreCase)
        .ToList();
}

public sealed record ImportEventRow(
    int EventId,
    string Name,
    string Type,
    string Cup,
    bool Selectable,
    bool Suggested,
    bool Counts);
