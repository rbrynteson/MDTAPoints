using System.Globalization;

namespace MDTAPoints.Core;

public sealed record WeekendFieldEntry(
    int TournamentId,
    string TournamentName,
    DateOnly Start,
    EventKind Event,
    Division Division,
    int EntryCount);

public sealed record WeekendDivisionAverage(EventKind Event, Division Division, double Average, int Weekends);

public sealed record WeekendField(
    string Week,
    DateOnly Start,
    DateOnly End,
    IReadOnlyList<string> Tournaments,
    IReadOnlyDictionary<(EventKind Event, Division Division), int> Counts)
{
    public int Count(EventKind kind, Division division) =>
        Counts.TryGetValue((kind, division), out var count) ? count : 0;

    public string Dates => WeekendFieldAnalytics.WeekendDates(Start, End);
}

public sealed record WeekendFieldReport(
    IReadOnlyList<WeekendField> Weekends,
    IReadOnlyList<WeekendDivisionAverage> Averages)
{
    public WeekendDivisionAverage Average(EventKind kind, Division division) =>
        Averages.Single(row => row.Event == kind && row.Division == division);
}

public static class WeekendFieldAnalytics
{
    public static readonly Division[] DivisionOrder =
    {
        Division.Novice,
        Division.JuniorVarsity,
        Division.Varsity
    };

    public static WeekendFieldReport Calculate(IEnumerable<WeekendFieldEntry> source)
    {
        var weekends = source
            .Where(entry => entry.EntryCount > 0)
            .GroupBy(entry => SeasonWeek.Key(entry.Start))
            .Select(SummarizeWeek)
            .OrderBy(weekend => weekend.Start)
            .ThenBy(weekend => weekend.Week, StringComparer.Ordinal)
            .ToList();

        var averages = new List<WeekendDivisionAverage>();
        foreach (var kind in ParticipationAnalytics.EventOrder)
        {
            foreach (var division in DivisionOrder)
            {
                var samples = weekends
                    .Select(weekend => weekend.Count(kind, division))
                    .Where(count => count > 0)
                    .ToList();
                averages.Add(new WeekendDivisionAverage(
                    kind,
                    division,
                    samples.Count == 0 ? 0 : samples.Average(),
                    samples.Count));
            }
        }

        return new WeekendFieldReport(weekends, averages);
    }

    public static string DivisionLabel(Division division) => division switch
    {
        Division.Novice => "Novice",
        Division.JuniorVarsity => "JV",
        Division.Varsity => "Varsity",
        _ => division.ToString()
    };

    public static string FormatAverage(double value)
    {
        if (value <= 0 || double.IsNaN(value))
        {
            return "—";
        }

        var rounded = Math.Round(value, 1, MidpointRounding.AwayFromZero);
        return Math.Abs(rounded % 1) < 0.05
            ? rounded.ToString("0", CultureInfo.InvariantCulture)
            : rounded.ToString("0.0", CultureInfo.InvariantCulture);
    }

    public static string FormatTick(double value) => FormatAverage(value) == "—"
        ? "0"
        : FormatAverage(value);

    public static string WeekendDates(DateOnly start, DateOnly end)
    {
        if (end <= start)
        {
            return start.ToString("MMM d", CultureInfo.InvariantCulture);
        }

        if (start.Year == end.Year && start.Month == end.Month)
        {
            return $"{start.ToString("MMM d", CultureInfo.InvariantCulture)}–{end.Day}";
        }

        return $"{start.ToString("MMM d", CultureInfo.InvariantCulture)}–{end.ToString("MMM d", CultureInfo.InvariantCulture)}";
    }

    public static string ChartLabel(WeekendFieldReport report)
    {
        var parts = ParticipationAnalytics.EventOrder.Select(kind =>
        {
            var bits = DivisionOrder.Select(division =>
            {
                var average = report.Average(kind, division);
                return DivisionLabel(division) + " " + FormatAverage(average.Average);
            });
            return EventKindLabels.Short(kind) + " " + string.Join(", ", bits);
        });
        return "Average entries per weekend. " + string.Join(". ", parts) + ".";
    }

    public static double ChartMax(double peak)
    {
        if (peak <= 0 || double.IsNaN(peak) || double.IsInfinity(peak))
        {
            return 1;
        }

        var padded = peak * 1.12;
        var exponent = Math.Floor(Math.Log10(padded));
        var magnitude = Math.Pow(10, exponent);
        var residual = padded / magnitude;
        var nice = residual <= 1 ? 1d
            : residual <= 1.2 ? 1.2
            : residual <= 1.5 ? 1.5
            : residual <= 2 ? 2
            : residual <= 2.5 ? 2.5
            : residual <= 3 ? 3
            : residual <= 4 ? 4
            : residual <= 5 ? 5
            : residual <= 6 ? 6
            : residual <= 8 ? 8
            : 10;
        return Math.Round(nice * magnitude, 8, MidpointRounding.AwayFromZero);
    }

    private static WeekendField SummarizeWeek(IGrouping<string, WeekendFieldEntry> group)
    {
        var counts = new Dictionary<(EventKind, Division), int>();
        foreach (var entry in group)
        {
            var key = (entry.Event, entry.Division);
            counts[key] = counts.GetValueOrDefault(key) + entry.EntryCount;
        }

        var tournaments = group
            .GroupBy(entry => entry.TournamentId)
            .Select(tournament => tournament.First().TournamentName)
            .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
            .ThenBy(name => name, StringComparer.Ordinal)
            .ToList();
        return new WeekendField(
            group.Key,
            group.Min(entry => entry.Start),
            group.Max(entry => entry.Start),
            tournaments,
            counts);
    }
}
