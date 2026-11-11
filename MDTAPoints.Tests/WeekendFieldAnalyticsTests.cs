using ClosedXML.Excel;
using MDTAPoints.Api.Services;
using MDTAPoints.Core;

namespace MDTAPoints.Tests;

public class WeekendFieldAnalyticsTests
{
    [Fact]
    public void Sites_on_the_same_weekend_add_and_open_counts_as_varsity()
    {
        var report = WeekendFieldAnalytics.Calculate(new[]
        {
            Field(1, "Rosemount", new DateOnly(2026, 10, 9), EventKind.LincolnDouglas, Division.Novice, 20),
            Field(1, "Rosemount", new DateOnly(2026, 10, 9), EventKind.LincolnDouglas, Division.JuniorVarsity, 10),
            Field(1, "Rosemount", new DateOnly(2026, 10, 9), EventKind.LincolnDouglas, Division.Varsity, 30),
            Field(2, "Coon Rapids", new DateOnly(2026, 10, 10), EventKind.LincolnDouglas, Division.Novice, 15),
            Field(2, "Coon Rapids", new DateOnly(2026, 10, 10), EventKind.LincolnDouglas, Division.Varsity, 25),
            Field(2, "Coon Rapids", new DateOnly(2026, 10, 10), EventKind.PublicForum, Division.Novice, 12),
            Field(1, "Rosemount", new DateOnly(2026, 10, 17), EventKind.LincolnDouglas, Division.Novice, 10),
            Field(1, "Rosemount", new DateOnly(2026, 10, 17), EventKind.LincolnDouglas, Division.Varsity, 40)
        });

        Assert.Equal(2, report.Weekends.Count);

        var combined = report.Weekends[0];
        Assert.Equal("Oct 9–10", combined.Dates);
        Assert.Equal(new[] { "Coon Rapids", "Rosemount" }, combined.Tournaments);
        Assert.Equal(35, combined.Count(EventKind.LincolnDouglas, Division.Novice));
        Assert.Equal(10, combined.Count(EventKind.LincolnDouglas, Division.JuniorVarsity));
        Assert.Equal(55, combined.Count(EventKind.LincolnDouglas, Division.Varsity));
        Assert.Equal(12, combined.Count(EventKind.PublicForum, Division.Novice));
        Assert.Equal(0, combined.Count(EventKind.PublicForum, Division.Varsity));

        var later = report.Weekends[1];
        Assert.Equal("Oct 17", later.Dates);
        Assert.Equal(new[] { "Rosemount" }, later.Tournaments);
        Assert.Equal(0, later.Count(EventKind.LincolnDouglas, Division.JuniorVarsity));

        var novice = report.Average(EventKind.LincolnDouglas, Division.Novice);
        Assert.Equal(22.5, novice.Average);
        Assert.Equal(2, novice.Weekends);

        var jv = report.Average(EventKind.LincolnDouglas, Division.JuniorVarsity);
        Assert.Equal(10, jv.Average);
        Assert.Equal(1, jv.Weekends);

        var varsity = report.Average(EventKind.LincolnDouglas, Division.Varsity);
        Assert.Equal(47.5, varsity.Average);
        Assert.Equal(2, varsity.Weekends);

        Assert.Equal(0, report.Average(EventKind.Policy, Division.Varsity).Weekends);
        Assert.Equal("22.5", WeekendFieldAnalytics.FormatAverage(novice.Average));
        Assert.Equal("10", WeekendFieldAnalytics.FormatAverage(jv.Average));
        Assert.Equal("—", WeekendFieldAnalytics.FormatAverage(0));
    }

    [Fact]
    public void Open_and_varsity_at_one_site_add_into_the_varsity_field()
    {
        var report = WeekendFieldAnalytics.Calculate(new[]
        {
            Field(1, "Rosemount", new DateOnly(2026, 10, 9), EventKind.LincolnDouglas, Division.Varsity, 30),
            Field(1, "Rosemount", new DateOnly(2026, 10, 9), EventKind.LincolnDouglas, Division.Varsity, 5),
            Field(2, "Coon Rapids", new DateOnly(2026, 10, 10), EventKind.LincolnDouglas, Division.Varsity, 25)
        });

        var weekend = report.Weekends.Single();
        Assert.Equal(60, weekend.Count(EventKind.LincolnDouglas, Division.Varsity));
        Assert.Equal(0, weekend.Count(EventKind.LincolnDouglas, Division.JuniorVarsity));
        Assert.Equal(60, report.Average(EventKind.LincolnDouglas, Division.Varsity).Average);
    }

    [Fact]
    public void An_empty_season_still_lists_every_event_and_division()
    {
        var report = WeekendFieldAnalytics.Calculate(Array.Empty<WeekendFieldEntry>());

        Assert.Empty(report.Weekends);
        Assert.Equal(ParticipationAnalytics.EventOrder.Length * WeekendFieldAnalytics.DivisionOrder.Length, report.Averages.Count);
        Assert.All(report.Averages, row => Assert.Equal(0, row.Weekends));
    }

    [Fact]
    public void Chart_scale_leaves_room_above_the_tallest_bar()
    {
        Assert.Equal(60, WeekendFieldAnalytics.ChartMax(47.5));
        Assert.Equal(20, WeekendFieldAnalytics.ChartMax(15));
        Assert.Equal(12, WeekendFieldAnalytics.ChartMax(9.9));
        Assert.Equal(12, WeekendFieldAnalytics.ChartMax(10));
        Assert.Equal(60, WeekendFieldAnalytics.ChartMax(50));
        Assert.Equal(1, WeekendFieldAnalytics.ChartMax(0));
    }

    [Fact]
    public void Chart_and_workbook_show_the_combined_weekend_field()
    {
        var weekends = WeekendFieldAnalytics.Calculate(new[]
        {
            Field(1, "Rosemount", new DateOnly(2026, 10, 9), EventKind.LincolnDouglas, Division.Varsity, 30),
            Field(2, "Coon Rapids", new DateOnly(2026, 10, 10), EventKind.LincolnDouglas, Division.Varsity, 25)
        });

        var svg = WeekendFieldChart.Render(weekends);
        Assert.Contains("role=\"img\"", svg);
        Assert.Contains("class=\"varsity\"", svg);
        Assert.Contains(">55<", svg);
        Assert.DoesNotContain("<script", svg, StringComparison.OrdinalIgnoreCase);

        var bytes = ExcelExports.Analytics(ParticipationAnalytics.Calculate(Array.Empty<ParticipationEntry>()), weekends);
        using var workbook = new XLWorkbook(new MemoryStream(bytes));
        var averages = workbook.Worksheet("Weekend averages");
        Assert.Equal("Lincoln-Douglas", averages.Cell(2, 1).GetString());
        Assert.Equal(55, averages.Cell(2, 6).GetValue<double>());
        Assert.Equal(1, averages.Cell(2, 7).GetValue<int>());

        var sheet = workbook.Worksheet("Weekends");
        Assert.Equal("Oct 9–10", sheet.Cell(2, 1).GetString());
        Assert.Equal("Coon Rapids, Rosemount", sheet.Cell(2, 2).GetString());
        Assert.Equal(55, sheet.Cell(2, 5).GetValue<int>());
    }

    private static WeekendFieldEntry Field(
        int tournamentId,
        string name,
        DateOnly start,
        EventKind kind,
        Division division,
        int entries) =>
        new(tournamentId, name, start, kind, division, entries);
}
