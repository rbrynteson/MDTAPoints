using System.Globalization;
using System.Net;
using System.Text;
using MDTAPoints.Core;

namespace MDTAPoints.Api.Services;

public static class WeekendFieldChart
{
    public static string Render(WeekendFieldReport report)
    {
        const double width = 760;
        const double height = 280;
        const double plotLeft = 52;
        const double plotTop = 36;
        const double plotWidth = 688;
        const double plotHeight = 198;
        var peak = report.Averages.Count == 0 ? 0 : report.Averages.Max(row => row.Average);
        var max = WeekendFieldAnalytics.ChartMax(peak);
        var events = ParticipationAnalytics.EventOrder;
        var divisions = WeekendFieldAnalytics.DivisionOrder;
        var groupWidth = plotWidth / events.Length;
        const double barWidth = 26;
        const double barGap = 8;
        var cluster = barWidth * divisions.Length + barGap * (divisions.Length - 1);

        var builder = new StringBuilder();
        builder.Append("<svg viewBox=\"0 0 ").Append(Num(width)).Append(' ').Append(Num(height))
            .Append("\" role=\"img\" aria-label=\"")
            .Append(Encode(WeekendFieldAnalytics.ChartLabel(report)))
            .Append("\">");

        for (var tick = 0; tick <= 2; tick++)
        {
            var value = max * tick / 2d;
            var y = plotTop + plotHeight - plotHeight * tick / 2d;
            Line(builder, plotLeft, y, plotLeft + plotWidth, y);
            Text(builder, "axis", plotLeft - 8, y + 4, "end", WeekendFieldAnalytics.FormatTick(value));
        }

        for (var eventIndex = 0; eventIndex < events.Length; eventIndex++)
        {
            var kind = events[eventIndex];
            var groupX = plotLeft + eventIndex * groupWidth;
            Text(builder, "axis", groupX + groupWidth / 2d, 258, "middle", EventKindLabels.Short(kind));

            for (var divisionIndex = 0; divisionIndex < divisions.Length; divisionIndex++)
            {
                var division = divisions[divisionIndex];
                var average = report.Average(kind, division);
                if (average.Weekends == 0 || average.Average <= 0)
                {
                    continue;
                }

                var barHeight = average.Average / max * plotHeight;
                var x = groupX + (groupWidth - cluster) / 2d + divisionIndex * (barWidth + barGap);
                var y = plotTop + plotHeight - barHeight;
                var css = division switch
                {
                    Division.Novice => "novice",
                    Division.JuniorVarsity => "jv",
                    _ => "varsity"
                };
                var label = EventKindLabels.Short(kind) + " " + WeekendFieldAnalytics.DivisionLabel(division) + ": " + WeekendFieldAnalytics.FormatAverage(average.Average);
                builder.Append("<rect class=\"").Append(css)
                    .Append("\" x=\"").Append(Num(x))
                    .Append("\" y=\"").Append(Num(y))
                    .Append("\" width=\"").Append(Num(barWidth))
                    .Append("\" height=\"").Append(Num(barHeight))
                    .Append("\"><title>").Append(Encode(label)).Append("</title></rect>");
                Text(builder, "bar-value", x + barWidth / 2d, y - 6, "middle", WeekendFieldAnalytics.FormatAverage(average.Average));
            }
        }

        builder.Append("</svg>");
        return builder.ToString();
    }

    private static void Line(StringBuilder builder, double x1, double y1, double x2, double y2)
    {
        builder.Append("<line class=\"grid\" x1=\"").Append(Num(x1))
            .Append("\" y1=\"").Append(Num(y1))
            .Append("\" x2=\"").Append(Num(x2))
            .Append("\" y2=\"").Append(Num(y2))
            .Append("\" />");
    }

    private static void Text(StringBuilder builder, string css, double x, double y, string anchor, string value)
    {
        builder.Append("<text class=\"").Append(css)
            .Append("\" x=\"").Append(Num(x))
            .Append("\" y=\"").Append(Num(y))
            .Append("\" text-anchor=\"").Append(anchor)
            .Append("\">").Append(Encode(value)).Append("</text>");
    }

    private static string Num(double value) => value.ToString("0.##", CultureInfo.InvariantCulture);

    private static string Encode(string value) => WebUtility.HtmlEncode(value);
}
