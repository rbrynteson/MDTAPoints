using MDTAPoints.Api.Data;
using MDTAPoints.Api.Services;
using MDTAPoints.Core;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;

namespace MDTAPoints.Api.Pages.Analytics;

public class IndexModel : PageModel
{
    private readonly CupDatabase _database;

    public IndexModel(CupDatabase database)
    {
        _database = database;
    }

    public ParticipationReport Report { get; private set; } = ParticipationAnalytics.Calculate(Array.Empty<ParticipationEntry>());
    public WeekendFieldReport Weekends { get; private set; } = WeekendFieldAnalytics.Calculate(Array.Empty<WeekendFieldEntry>());
    public bool AllDivisions { get; private set; }
    public string ChartSvg => WeekendFieldChart.Render(Weekends);

    public async Task OnGetAsync(string? divisions)
    {
        AllDivisions = IsAll(divisions);
        Report = ParticipationAnalytics.Calculate(await _database.LoadParticipationAsync(AllDivisions, HttpContext.RequestAborted));
        Weekends = WeekendFieldAnalytics.Calculate(await _database.LoadWeekendFieldsAsync(HttpContext.RequestAborted));
    }

    public async Task<IActionResult> OnGetExportAsync(string? divisions)
    {
        var all = IsAll(divisions);
        var report = ParticipationAnalytics.Calculate(await _database.LoadParticipationAsync(all, HttpContext.RequestAborted));
        var weekends = WeekendFieldAnalytics.Calculate(await _database.LoadWeekendFieldsAsync(HttpContext.RequestAborted));
        var fileName = all ? "mdta-cup-analytics-all-divisions.xlsx" : "mdta-cup-analytics.xlsx";
        return File(ExcelExports.Analytics(report, weekends), ExcelExports.ContentType, fileName);
    }

    private static bool IsAll(string? divisions) =>
        string.Equals(divisions, "all", StringComparison.OrdinalIgnoreCase);
}
