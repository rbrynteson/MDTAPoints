using System.Text;
using ClosedXML.Excel;
using MDTAPoints.Api.Data;
using MDTAPoints.Core;

namespace MDTAPoints.Api.Services;

public static class ExcelExports
{
    public const string ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

    public static byte[] Standings(
        IReadOnlyList<IndividualStanding> individuals,
        IReadOnlyList<TeamStanding> teams,
        bool includeIndividuals,
        bool includeTeams,
        bool includeEvent)
    {
        if (!includeIndividuals && !includeTeams)
        {
            includeIndividuals = true;
            includeTeams = true;
        }

        using var workbook = new XLWorkbook();
        if (includeIndividuals)
        {
            WriteIndividuals(workbook, individuals, includeEvent);
        }

        if (includeTeams)
        {
            WriteTeams(workbook, teams, includeEvent);
        }

        return Save(workbook);
    }

    public static byte[] Students(IReadOnlyList<StudentAdminRow> students, IReadOnlyList<HeadToHeadRow> headToHead)
    {
        using var workbook = new XLWorkbook();
        var events = EventOrder();
        var headers = new List<string> { "Id", "Student", "School", "Grade" };
        foreach (var kind in events)
        {
            var label = EventKindLabels.Display(kind);
            headers.Add(label + " manual");
            headers.Add(label + " section");
        }

        var sheet = AddSheet(workbook, "Students", headers);
        for (var index = 0; index < students.Count; index++)
        {
            var student = students[index];
            var row = index + 2;
            var column = 1;
            sheet.Cell(row, column++).Value = student.Id;
            sheet.Cell(row, column++).Value = student.DisplayName;
            sheet.Cell(row, column++).Value = student.SchoolName;
            if (student.Grade is int grade)
            {
                sheet.Cell(row, column).Value = grade;
            }

            column++;
            foreach (var kind in events)
            {
                var key = kind.ToString();
                sheet.Cell(row, column++).Value = student.ManualEvents.Contains(key) ? "Yes" : "";
                sheet.Cell(row, column++).Value = student.SectionEvents.Contains(key) ? "Yes" : "";
            }
        }

        Finish(sheet);

        var matches = AddSheet(workbook, "Head-to-head", new[] { "Winner id", "Winner", "Loser id", "Loser" });
        for (var index = 0; index < headToHead.Count; index++)
        {
            var result = headToHead[index];
            var row = index + 2;
            matches.Cell(row, 1).Value = result.WinnerStudentId;
            matches.Cell(row, 2).Value = result.WinnerName;
            matches.Cell(row, 3).Value = result.LoserStudentId;
            matches.Cell(row, 4).Value = result.LoserName;
        }

        Finish(matches);
        return Save(workbook);
    }

    public static byte[] Tournaments(IReadOnlyList<TournamentListItem> tournaments)
    {
        using var workbook = new XLWorkbook();
        var sheet = AddSheet(workbook, "Tournaments", new[] { "Date", "Name", "Type", "Status" });
        for (var index = 0; index < tournaments.Count; index++)
        {
            var tournament = tournaments[index];
            var row = index + 2;
            var date = sheet.Cell(row, 1);
            date.Value = tournament.StartDate.ToDateTime(TimeOnly.MinValue);
            date.Style.NumberFormat.Format = "yyyy-mm-dd";
            sheet.Cell(row, 2).Value = tournament.Name;
            sheet.Cell(row, 3).Value = tournament.Kind;
            sheet.Cell(row, 4).Value = tournament.Status;
        }

        Finish(sheet);
        return Save(workbook);
    }

    public static byte[] Users(IReadOnlyList<AppUser> users)
    {
        using var workbook = new XLWorkbook();
        var sheet = AddSheet(workbook, "Users", new[] { "Name", "Username" });
        for (var index = 0; index < users.Count; index++)
        {
            var user = users[index];
            var row = index + 2;
            sheet.Cell(row, 1).Value = user.DisplayName;
            sheet.Cell(row, 2).Value = user.Username;
        }

        Finish(sheet);
        return Save(workbook);
    }

    public static byte[] Tournament(TournamentDetail tournament)
    {
        using var workbook = new XLWorkbook();
        if (tournament.Events.Count == 0)
        {
            Finish(AddSheet(workbook, "Entries", EntryHeaders()));
            return Save(workbook);
        }

        foreach (var item in tournament.Events)
        {
            var sheet = AddSheet(workbook, string.IsNullOrWhiteSpace(item.Name) ? item.EventKind : item.Name, EntryHeaders());
            for (var index = 0; index < item.Entries.Count; index++)
            {
                var entry = item.Entries[index];
                var row = index + 2;
                sheet.Cell(row, 1).Value = entry.Place;
                sheet.Cell(row, 2).Value = entry.EntryName;
                sheet.Cell(row, 3).Value = entry.SchoolName;
                SetMultiline(sheet.Cell(row, 4), string.Join("\n", entry.Competitors.Select(competitor => $"{competitor.LinkStatus}: {competitor.RawName}")));
                sheet.Cell(row, 5).Value = entry.ElimWins;
                sheet.Cell(row, 6).Value = Flags(entry);
                if (entry.TabroomPlace is int posted)
                {
                    sheet.Cell(row, 7).Value = posted;
                }

                sheet.Cell(row, 8).Value = entry.PlaceNote;
            }

            Finish(sheet);
        }

        return Save(workbook);
    }

    public static byte[] EventResults(string eventName, IReadOnlyList<TournamentResultRow> rows)
    {
        using var workbook = new XLWorkbook();
        var sheet = AddSheet(workbook, string.IsNullOrWhiteSpace(eventName) ? "Results" : eventName, new[] { "Place", "Name", "School", "Points", "Scoring", "Place note" });
        for (var index = 0; index < rows.Count; index++)
        {
            var row = rows[index];
            var line = index + 2;
            sheet.Cell(line, 1).Value = row.Place;
            sheet.Cell(line, 2).Value = row.Name;
            sheet.Cell(line, 3).Value = row.SchoolName;
            sheet.Cell(line, 4).Value = row.Points;
            sheet.Cell(line, 5).Value = row.Scoring;
            sheet.Cell(line, 6).Value = row.PlaceNote;
        }

        Finish(sheet);
        return Save(workbook);
    }

    public static string FileSlug(string value)
    {
        var builder = new StringBuilder();
        var pendingHyphen = false;
        foreach (var ch in value.Trim().ToLowerInvariant())
        {
            if (char.IsAsciiLetterOrDigit(ch))
            {
                if (pendingHyphen && builder.Length > 0)
                {
                    builder.Append('-');
                }

                pendingHyphen = false;
                builder.Append(ch);
            }
            else
            {
                pendingHyphen = true;
            }
        }

        return builder.Length == 0 ? "export" : builder.ToString();
    }

    private static void WriteIndividuals(XLWorkbook workbook, IReadOnlyList<IndividualStanding> rows, bool includeEvent)
    {
        var headers = TableHeaders(includeEvent, "Award", "Student", "School", "Points", "Tournaments", "Breakdown");
        var sheet = AddSheet(workbook, "Individuals", headers);
        for (var index = 0; index < rows.Count; index++)
        {
            var standing = rows[index];
            var row = index + 2;
            var column = WriteEvent(sheet, row, standing.Event, includeEvent);
            if (!standing.Eligible)
            {
                sheet.Cell(row, column).Value = "Not eligible";
            }
            else if (standing.AwardRank is int rank)
            {
                sheet.Cell(row, column).Value = rank;
            }

            column++;
            sheet.Cell(row, column++).Value = standing.Name;
            sheet.Cell(row, column++).Value = standing.SchoolName;
            sheet.Cell(row, column++).Value = standing.Points;
            sheet.Cell(row, column++).Value = standing.TournamentCount;
            SetMultiline(sheet.Cell(row, column), FormatLines(standing.Lines, individual: true));
        }

        Finish(sheet);
    }

    private static void WriteTeams(XLWorkbook workbook, IReadOnlyList<TeamStanding> rows, bool includeEvent)
    {
        var headers = TableHeaders(includeEvent, "Place", "School", "Points", "Tournaments", "Breakdown");
        var sheet = AddSheet(workbook, "Teams", headers);
        for (var index = 0; index < rows.Count; index++)
        {
            var standing = rows[index];
            var row = index + 2;
            var column = WriteEvent(sheet, row, standing.Event, includeEvent);
            sheet.Cell(row, column++).Value = standing.Rank;
            sheet.Cell(row, column++).Value = standing.SchoolName;
            sheet.Cell(row, column++).Value = standing.Points;
            sheet.Cell(row, column++).Value = standing.TournamentCount;
            SetMultiline(sheet.Cell(row, column), FormatLines(standing.Lines, individual: false));
        }

        Finish(sheet);
    }

    private static IReadOnlyList<string> TableHeaders(bool includeEvent, params string[] columns)
    {
        if (!includeEvent)
        {
            return columns;
        }

        var headers = new List<string> { "Event" };
        headers.AddRange(columns);
        return headers;
    }

    private static int WriteEvent(IXLWorksheet sheet, int row, EventKind kind, bool includeEvent)
    {
        if (!includeEvent)
        {
            return 1;
        }

        sheet.Cell(row, 1).Value = EventKindLabels.Display(kind);
        return 2;
    }

    private static string FormatLines(IReadOnlyList<PointLine> lines, bool individual)
    {
        return string.Join("\n", lines.Select(line =>
        {
            if (line.Dropped)
            {
                return individual
                    ? $"Dropped: {line.TournamentName} ({line.Points})"
                    : $"Dropped: {line.TournamentName}, {line.Points} pts";
            }

            return individual
                ? $"{line.Detail}: {line.TournamentName} ({line.Points})"
                : $"Kept: {line.TournamentName}, {line.Points} pts";
        }));
    }

    private static string Flags(EntryRow entry)
    {
        var flags = new List<string>();
        if (entry.ReachedElims)
        {
            flags.Add("Elims");
        }

        if (entry.ReachedFinals)
        {
            flags.Add("Finals");
        }

        if (entry.ReachedSemis)
        {
            flags.Add("Semis");
        }

        if (entry.QualifiedToState)
        {
            flags.Add("Qualified to state");
        }

        return string.Join(", ", flags);
    }

    private static string[] EntryHeaders() => new[] { "Place", "Entry", "School", "Names", "Elim wins", "Flags", "Tabroom place", "Place note" };

    private static EventKind[] EventOrder() => new[]
    {
        EventKind.LincolnDouglas,
        EventKind.PublicForum,
        EventKind.Policy,
        EventKind.Classic,
        EventKind.Congress
    };

    private static IXLWorksheet AddSheet(XLWorkbook workbook, string name, IReadOnlyList<string> headers)
    {
        var sheet = workbook.Worksheets.Add(UniqueSheetName(workbook, name));
        for (var column = 0; column < headers.Count; column++)
        {
            var cell = sheet.Cell(1, column + 1);
            cell.Value = headers[column];
            cell.Style.Font.Bold = true;
            cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#E7EEF5");
        }

        sheet.SheetView.FreezeRows(1);
        sheet.Row(1).Height = 18;
        return sheet;
    }

    private static void Finish(IXLWorksheet sheet)
    {
        var lastColumn = sheet.LastColumnUsed()?.ColumnNumber() ?? 1;
        var lastRow = Math.Max(1, sheet.LastRowUsed()?.RowNumber() ?? 1);
        sheet.Range(1, 1, lastRow, lastColumn).SetAutoFilter();
        if (lastRow >= 2)
        {
            sheet.Range(2, 1, lastRow, lastColumn).Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
        }

        for (var column = 1; column <= lastColumn; column++)
        {
            sheet.Column(column).Width = sheet.Cell(1, column).GetString() switch
            {
                "Award" or "Place" or "Id" or "Grade" or "Points" or "Elim wins" or "Winner id" or "Loser id" => 12,
                "Tournaments" or "Type" or "Status" or "Date" => 16,
                "Student" or "School" or "Name" or "Entry" or "Username" or "Winner" or "Loser" => 28,
                "Breakdown" or "Names" or "Scoring" => 56,
                "Place note" => 32,
                "Flags" => 32,
                "Event" => 24,
                _ => 22
            };
        }
    }

    private static void SetMultiline(IXLCell cell, string text)
    {
        cell.Value = text;
        cell.Style.Alignment.WrapText = true;
        var lines = text.Count(ch => ch == '\n') + 1;
        if (lines > 1)
        {
            var height = Math.Min(15d * lines, 75d);
            if (cell.WorksheetRow().Height < height)
            {
                cell.WorksheetRow().Height = height;
            }
        }
    }

    private static string UniqueSheetName(XLWorkbook workbook, string name)
    {
        var invalid = new HashSet<char> { ':', '\\', '/', '?', '*', '[', ']' };
        var builder = new StringBuilder();
        var pendingSpace = false;
        foreach (var ch in name)
        {
            if (invalid.Contains(ch) || char.IsWhiteSpace(ch))
            {
                pendingSpace = true;
                continue;
            }

            if (ch == '\'')
            {
                continue;
            }

            if (pendingSpace && builder.Length > 0)
            {
                builder.Append(' ');
            }

            pendingSpace = false;
            builder.Append(ch);
        }

        var cleaned = builder.ToString();
        if (cleaned.Length == 0)
        {
            cleaned = "Sheet";
        }

        if (cleaned.Length > 31)
        {
            cleaned = cleaned[..31].Trim();
        }

        var candidate = cleaned;
        var suffix = 2;
        while (workbook.Worksheets.Any(sheet => string.Equals(sheet.Name, candidate, StringComparison.OrdinalIgnoreCase)))
        {
            var tail = " " + suffix;
            var stemLength = Math.Max(1, Math.Min(cleaned.Length, 31 - tail.Length));
            candidate = cleaned[..stemLength].Trim() + tail;
            suffix++;
        }

        return candidate;
    }

    private static byte[] Save(XLWorkbook workbook)
    {
        using var stream = new MemoryStream();
        workbook.SaveAs(stream);
        return stream.ToArray();
    }
}
