using ClosedXML.Excel;
using MDTAPoints.Api.Data;
using MDTAPoints.Api.Services;
using MDTAPoints.Core;

namespace MDTAPoints.Tests;

public class ExcelExportTests
{
    [Fact]
    public void Standings_workbook_matches_the_individual_and_team_tables()
    {
        var bytes = ExcelExports.Standings(
            new[] { Individual(eligible: false, rank: null, points: 12) },
            new[] { Team() },
            includeIndividuals: true,
            includeTeams: true,
            includeEvent: false);

        using var workbook = Open(bytes);
        var individuals = workbook.Worksheet("Individuals");
        Assert.Equal(new[] { "Award", "Student", "School", "Points", "Tournaments", "Breakdown" }, Headers(individuals));
        Assert.Equal("Not eligible", individuals.Cell(2, 1).GetString());
        Assert.Equal("Ada Lovelace", individuals.Cell(2, 2).GetString());
        Assert.Equal(12, individuals.Cell(2, 4).GetValue<int>());
        Assert.Equal("Kept: Edina (8)\nDropped: Minnetonka (4)", individuals.Cell(2, 6).GetString());

        var teams = workbook.Worksheet("Teams");
        Assert.Equal(1, teams.Cell(2, 1).GetValue<int>());
        Assert.Equal("Edina", teams.Cell(2, 2).GetString());
        Assert.Equal("Kept: Edina, 8 pts\nDropped: Minnetonka, 4 pts", teams.Cell(2, 5).GetString());
    }

    [Fact]
    public void All_events_export_adds_the_event_column_and_keeps_award_ranks_numeric()
    {
        var bytes = ExcelExports.Standings(
            new[] { Individual(eligible: true, rank: 1, points: 20) },
            Array.Empty<TeamStanding>(),
            includeIndividuals: true,
            includeTeams: false,
            includeEvent: true);

        using var workbook = Open(bytes);
        Assert.Single(workbook.Worksheets);
        var sheet = workbook.Worksheet("Individuals");
        Assert.Equal("Event", sheet.Cell(1, 1).GetString());
        Assert.Equal("Lincoln-Douglas", sheet.Cell(2, 1).GetString());
        Assert.Equal(1, sheet.Cell(2, 2).GetValue<int>());
    }

    [Fact]
    public void Student_export_records_qualifier_marks_and_head_to_head()
    {
        var student = new StudentAdminRow
        {
            Id = 4,
            DisplayName = "Ada Lovelace",
            SchoolName = "Edina",
            Grade = 12,
            ManualEvents = new HashSet<string> { nameof(EventKind.LincolnDouglas) },
            SectionEvents = new HashSet<string> { nameof(EventKind.LincolnDouglas), nameof(EventKind.Policy) }
        };
        var bytes = ExcelExports.Students(
            new[] { student },
            new[] { new HeadToHeadRow { WinnerStudentId = 4, WinnerName = "Ada Lovelace", LoserStudentId = 9, LoserName = "Grace Hopper" } });

        using var workbook = Open(bytes);
        var students = workbook.Worksheet("Students");
        var headers = Headers(students);
        Assert.Equal("Yes", students.Cell(2, headers.IndexOf("Lincoln-Douglas manual") + 1).GetString());
        Assert.Equal("Yes", students.Cell(2, headers.IndexOf("Policy section") + 1).GetString());
        Assert.Equal("", students.Cell(2, headers.IndexOf("Policy manual") + 1).GetString());
        Assert.Equal(12, students.Cell(2, 4).GetValue<int>());

        var matches = workbook.Worksheet("Head-to-head");
        Assert.Equal("Grace Hopper", matches.Cell(2, 4).GetString());
    }

    [Fact]
    public void User_export_leaves_out_the_password_hash()
    {
        var bytes = ExcelExports.Users(new[]
        {
            new AppUser { Id = 1, DisplayName = "Ada", Username = "ada", PasswordHash = "secret-hash" }
        });

        using var workbook = Open(bytes);
        Assert.Equal(new[] { "Name", "Username" }, Headers(workbook.Worksheet("Users")));
        Assert.DoesNotContain(workbook.Worksheets.SelectMany(sheet => sheet.CellsUsed()), cell => cell.GetString().Contains("secret-hash"));
    }

    [Fact]
    public void Tournament_export_uses_one_sheet_per_event_table()
    {
        var tournament = new TournamentDetail
        {
            Name = "Edina Invitational",
            Events = new List<EventRow>
            {
                new()
                {
                    Name = "LD / Varsity",
                    Entries = new List<EntryRow>
                    {
                        new()
                        {
                            Place = 1,
                            TabroomPlace = 4,
                            PlaceNote = "Hidden semis",
                            EntryName = "Ada Lovelace",
                            SchoolName = "Edina",
                            ElimWins = 3,
                            ReachedFinals = true,
                            QualifiedToState = true,
                            Competitors = new List<CompetitorRow>
                            {
                                new() { LinkStatus = "Linked", RawName = "Ada Lovelace" }
                            }
                        }
                    }
                },
                new() { Name = "LD / Varsity" }
            }
        };

        var bytes = ExcelExports.Tournament(tournament);
        using var workbook = Open(bytes);
        Assert.Equal(new[] { "LD Varsity", "LD Varsity 2" }, workbook.Worksheets.Select(sheet => sheet.Name).ToArray());
        var sheet = workbook.Worksheet("LD Varsity");
        Assert.Equal("Linked: Ada Lovelace", sheet.Cell(2, 4).GetString());
        Assert.Equal("Finals, Qualified to state", sheet.Cell(2, 6).GetString());
        Assert.Equal(4, sheet.Cell(2, 7).GetValue<int>());
        Assert.Equal("Hidden semis", sheet.Cell(2, 8).GetString());
        Assert.Equal("edina-invitational", ExcelExports.FileSlug(tournament.Name));
    }

    [Fact]
    public void Event_results_export_lists_place_name_and_points()
    {
        var bytes = ExcelExports.EventResults("Varsity LD", new[]
        {
            new TournamentResultRow(1, "Ada Lovelace", "Edina", 4, "Competed, Placed first", "Hidden semis")
        });

        using var workbook = Open(bytes);
        var sheet = workbook.Worksheet("Varsity LD");
        Assert.Equal(new[] { "Place", "Name", "School", "Points", "Scoring", "Place note" }, Headers(sheet));
        Assert.Equal(1, sheet.Cell(2, 1).GetValue<int>());
        Assert.Equal("Ada Lovelace", sheet.Cell(2, 2).GetString());
        Assert.Equal(4, sheet.Cell(2, 4).GetValue<int>());
        Assert.Equal("Competed, Placed first", sheet.Cell(2, 5).GetString());
        Assert.Equal("Hidden semis", sheet.Cell(2, 6).GetString());
    }

    private static IndividualStanding Individual(bool eligible, int? rank, int points)
    {
        return new IndividualStanding(
            "ada",
            "Ada Lovelace",
            "edina",
            "Edina",
            EventKind.LincolnDouglas,
            points,
            eligible,
            2,
            false,
            11,
            4.5,
            rank,
            new[]
            {
                new PointLine("t1", "Edina", new DateOnly(2026, 10, 3), 8, 3, false, "Kept"),
                new PointLine("t2", "Minnetonka", new DateOnly(2026, 10, 4), 4, 8, true, "Kept")
            });
    }

    private static TeamStanding Team()
    {
        return new TeamStanding(
            "edina",
            "Edina",
            EventKind.PublicForum,
            20,
            1,
            3,
            2.5,
            new[]
            {
                new PointLine("t1", "Edina", new DateOnly(2026, 10, 3), 8, 1, false, "Kept"),
                new PointLine("t2", "Minnetonka", new DateOnly(2026, 10, 4), 4, 4, true, "Kept")
            });
    }

    private static XLWorkbook Open(byte[] bytes) => new(new MemoryStream(bytes));

    private static List<string> Headers(IXLWorksheet sheet)
    {
        var last = sheet.LastColumnUsed()?.ColumnNumber() ?? 0;
        return Enumerable.Range(1, last).Select(column => sheet.Cell(1, column).GetString()).ToList();
    }
}
