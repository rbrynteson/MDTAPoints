using MDTAPoints.Core;

namespace MDTAPoints.Tests;

public class StandingsCalculatorTests
{
    [Fact]
    public void A_student_keeps_the_higher_tournament_in_the_same_week()
    {
        var standings = StandingsCalculator.Calculate(new SeasonInput(
            new[]
            {
                DebateEntry("a", new DateOnly(2026, 10, 10), "Ada", "edina", place: 1, pointsHint: 0),
                DebateEntry("b", new DateOnly(2026, 10, 11), "Ada", "edina", place: 20, pointsHint: 0)
            },
            Array.Empty<ManualQualification>(),
            Array.Empty<HeadToHeadResult>()));

        var ada = standings.Individuals.Single();
        Assert.Equal(4, ada.Points);
        Assert.Contains(ada.Lines, line => line.Dropped && line.TournamentId == "b");
        Assert.False(ada.Eligible);
    }

    [Fact]
    public void Four_in_state_invitationals_and_a_state_qualifier_make_the_award_total()
    {
        var entries = Enumerable.Range(1, 4)
            .Select(week => DebateEntry("t" + week, new DateOnly(2026, 9, 5).AddDays(7 * week), "Ada", "edina", place: 13))
            .ToList();
        var standings = StandingsCalculator.Calculate(new SeasonInput(
            entries,
            new[] { new ManualQualification("ada", EventKind.LincolnDouglas, true) },
            Array.Empty<HeadToHeadResult>()));

        var ada = standings.Individuals.Single();
        Assert.Equal(5, ada.Points);
        Assert.True(ada.Eligible);
        Assert.Equal(1, ada.AwardRank);
        Assert.True(ada.QualifiedToState);
    }

    [Fact]
    public void An_in_state_nsda_congress_result_counts_toward_the_four_tournament_minimum()
    {
        var entries = Enumerable.Range(1, 3)
            .Select(week => CongressEntry("t" + week, new DateOnly(2026, 9, 5).AddDays(7 * week), TournamentKind.Invitational))
            .Append(CongressEntry("nsda", new DateOnly(2026, 11, 7), TournamentKind.NsdaQualifier))
            .ToList();

        var ada = StandingsCalculator.Calculate(new SeasonInput(entries, Array.Empty<ManualQualification>(), Array.Empty<HeadToHeadResult>()))
            .Individuals.Single();

        Assert.True(ada.Eligible);
        Assert.Equal(4, ada.TournamentCount);
    }

    [Fact]
    public void Tie_breaks_follow_state_qualification_grade_head_to_head_then_schedule()
    {
        var entries = new List<SeasonEntry>
        {
            DebateEntry("t1", new DateOnly(2026, 9, 12), "Senior", "edina", place: 13, studentId: "senior", grade: 12),
            DebateEntry("t1", new DateOnly(2026, 9, 12), "Junior", "eden prairie", place: 13, studentId: "junior", grade: 11),
            DebateEntry("t2", new DateOnly(2026, 9, 19), "Qualified", "minnetonka", place: 13, studentId: "qualified", grade: 10),
            DebateEntry("t2", new DateOnly(2026, 9, 19), "Plain", "chanhassen", place: 13, studentId: "plain", grade: 10)
        };
        for (var week = 3; week <= 4; week++)
        {
            var date = new DateOnly(2026, 9, 5).AddDays(7 * week);
            entries.Add(DebateEntry("w" + week, date, "Senior", "edina", place: 13, studentId: "senior", grade: 12));
            entries.Add(DebateEntry("w" + week, date, "Junior", "eden prairie", place: 13, studentId: "junior", grade: 11));
            entries.Add(DebateEntry("w" + week, date, "Qualified", "minnetonka", place: 13, studentId: "qualified", grade: 10));
            entries.Add(DebateEntry("w" + week, date, "Plain", "chanhassen", place: 13, studentId: "plain", grade: 10));
        }

        var rows = StandingsCalculator.Calculate(new SeasonInput(
            entries,
            new[] { new ManualQualification("qualified", EventKind.LincolnDouglas, true) },
            Array.Empty<HeadToHeadResult>())).Individuals;

        Assert.Equal(new[] { "Qualified", "Senior", "Junior", "Plain" }, rows.Select(row => row.Name).ToArray());
    }

    [Fact]
    public void Head_to_head_applies_only_to_a_two_person_tie()
    {
        var entries = new List<SeasonEntry>();
        foreach (var student in new[] { ("ada", "Ada"), ("bea", "Bea"), ("cy", "Cy") })
        {
            for (var week = 1; week <= 4; week++)
            {
                entries.Add(DebateEntry("t" + week, new DateOnly(2026, 9, 5).AddDays(7 * week), student.Item2, "edina", place: 13, studentId: student.Item1, grade: 12));
            }
        }

        var withHeadToHead = StandingsCalculator.Calculate(new SeasonInput(
            entries.Where(entry => entry.Competitors[0].StudentId != "cy").ToList(),
            Array.Empty<ManualQualification>(),
            new[] { new HeadToHeadResult("bea", "ada") })).Individuals;
        var threeWay = StandingsCalculator.Calculate(new SeasonInput(
            entries,
            Array.Empty<ManualQualification>(),
            new[] { new HeadToHeadResult("bea", "ada") })).Individuals;

        Assert.Equal(new[] { "Bea", "Ada" }, withHeadToHead.Select(row => row.Name).ToArray());
        Assert.Equal(new[] { "Ada", "Bea", "Cy" }, threeWay.Select(row => row.Name).ToArray());
    }

    [Fact]
    public void Team_points_count_a_partnership_once_and_keep_one_tournament_per_week()
    {
        var weekend = new DateOnly(2026, 10, 10);
        var standings = StandingsCalculator.Calculate(new SeasonInput(
            new[]
            {
                Partnership("a", weekend, "entry-1", 1, "ada", "bea"),
                Partnership("a", weekend, "entry-2", 6, "cy", "dee"),
                Partnership("a", weekend, "entry-3", 20, "ev", "fay"),
                Partnership("b", weekend.AddDays(1), "entry-4", 1, "ada", "bea")
            },
            Array.Empty<ManualQualification>(),
            Array.Empty<HeadToHeadResult>()));

        var school = standings.Teams.Single(team => team.Event == EventKind.PublicForum);
        Assert.Equal(7, school.Points);
        Assert.Contains(school.Lines, line => line.Dropped);
        Assert.Equal(4, standings.Individuals.Single(row => row.Name == "ada").Points);
        Assert.Equal(3, standings.Individuals.Single(row => row.Name == "cy").Points);
    }

    [Fact]
    public void Each_event_keeps_its_own_team_winner()
    {
        var entries = new List<SeasonEntry>
        {
            DebateEntry("ld", new DateOnly(2026, 10, 10), "A One", "alpha", place: 1, schoolId: "alpha"),
            DebateEntry("ld", new DateOnly(2026, 10, 10), "C One", "charlie", place: 2, schoolId: "charlie"),
            DebateEntry("ld", new DateOnly(2026, 10, 10), "B One", "bravo", place: 6, schoolId: "bravo"),
            Partnership("pf", new DateOnly(2026, 10, 17), "alpha-pf", 1, "A Two", "A Three", schoolId: "alpha"),
            Partnership("pf", new DateOnly(2026, 10, 17), "bravo-pf", 6, "B Two", "B Three", schoolId: "bravo")
        };

        var standings = StandingsCalculator.Calculate(new SeasonInput(entries, Array.Empty<ManualQualification>(), Array.Empty<HeadToHeadResult>()));
        var ld = standings.Teams.Where(team => team.Event == EventKind.LincolnDouglas).OrderBy(team => team.Rank).ToList();
        var pf = standings.Teams.Where(team => team.Event == EventKind.PublicForum).OrderBy(team => team.Rank).ToList();

        Assert.Equal(new[] { "alpha", "charlie", "bravo" }, ld.Select(team => team.SchoolId).ToArray());
        Assert.Equal(new[] { 1, 2, 3 }, ld.Select(team => team.Rank).ToArray());
        Assert.Equal(new[] { "alpha", "bravo" }, pf.Select(team => team.SchoolId).ToArray());
        Assert.Equal(new[] { 1, 2 }, pf.Select(team => team.Rank).ToArray());
    }

    private static SeasonEntry DebateEntry(
        string tournamentId,
        DateOnly date,
        string name,
        string school,
        int place,
        int pointsHint = 0,
        string? studentId = null,
        int? grade = null,
        string? schoolId = null)
    {
        _ = pointsHint;
        var id = studentId ?? name.ToLowerInvariant();
        var schoolKey = schoolId ?? school;
        return new SeasonEntry(
            tournamentId,
            tournamentId,
            date,
            true,
            TournamentKind.Invitational,
            EventKind.LincolnDouglas,
            true,
            8,
            40,
            false,
            false,
            false,
            false,
            tournamentId + ":" + id,
            schoolKey,
            school,
            place,
            false,
            0,
            false,
            false,
            new[] { new SeasonCompetitor(id, name, schoolKey, school, grade) });
    }

    private static SeasonEntry CongressEntry(string tournamentId, DateOnly date, TournamentKind kind) =>
        DebateEntry(tournamentId, date, "Ada", "edina", place: 12, studentId: "ada") with
        {
            Event = EventKind.Congress,
            TournamentKind = kind,
            EntryCount = 20,
            HasFinalSession = true,
            ReachedFinals = false,
            IsHouse = false
        };

    private static SeasonEntry Partnership(
        string tournamentId,
        DateOnly date,
        string entryId,
        int place,
        string first,
        string second,
        string schoolId = "edina") =>
        new(
            tournamentId,
            tournamentId,
            date,
            true,
            TournamentKind.Invitational,
            EventKind.PublicForum,
            true,
            8,
            40,
            false,
            false,
            false,
            false,
            entryId,
            schoolId,
            schoolId,
            place,
            false,
            0,
            false,
            false,
            new[]
            {
                new SeasonCompetitor(first.ToLowerInvariant(), first, schoolId, schoolId, null),
                new SeasonCompetitor(second.ToLowerInvariant(), second, schoolId, schoolId, null)
            });
}
