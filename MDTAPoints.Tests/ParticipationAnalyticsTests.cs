using MDTAPoints.Core;

namespace MDTAPoints.Tests;

public class ParticipationAnalyticsTests
{
    [Fact]
    public void Individual_events_count_each_student_once()
    {
        var report = ParticipationAnalytics.Calculate(new[]
        {
            Entry(EventKind.LincolnDouglas, 1, 10, "Apple Valley", (5, "jane doe")),
            Entry(EventKind.LincolnDouglas, 2, 10, "Apple Valley", (5, "jane doe")),
            Entry(EventKind.LincolnDouglas, 3, 10, "Apple Valley", (6, "alex kim")),
            Entry(EventKind.Congress, 4, 11, "Edina", (7, "sam lee"))
        });

        var ld = EventRow(report, EventKind.LincolnDouglas);
        Assert.Equal(1, ld.SchoolCount);
        Assert.Equal(2, ld.Count);
        Assert.False(ld.CountEntries);

        var congress = EventRow(report, EventKind.Congress);
        Assert.Equal(1, congress.SchoolCount);
        Assert.Equal(1, congress.Count);
        Assert.False(congress.CountEntries);

        var apple = report.Schools.Single(school => school.SchoolId == 10);
        Assert.Equal(2, apple.Counts[EventKind.LincolnDouglas]);
        Assert.Equal(0, apple.Counts[EventKind.Congress]);
        Assert.Equal(2, apple.Total);
    }

    [Fact]
    public void Team_events_count_entries_and_keep_a_partnership_together()
    {
        var report = ParticipationAnalytics.Calculate(new[]
        {
            Entry(EventKind.PublicForum, 1, 10, "Apple Valley", (1, "a"), (2, "b")),
            Entry(EventKind.PublicForum, 2, 10, "Apple Valley", (2, "b"), (1, "a")),
            Entry(EventKind.PublicForum, 3, 10, "Apple Valley", (1, "a"), (3, "c")),
            Entry(EventKind.Policy, 4, 12, "Wayzata", (8, "d"), (9, "e")),
            Entry(EventKind.Classic, 5, 12, "Wayzata", (8, "d"), (9, "e"))
        });

        var pf = EventRow(report, EventKind.PublicForum);
        Assert.Equal(1, pf.SchoolCount);
        Assert.Equal(2, pf.Count);
        Assert.True(pf.CountEntries);

        Assert.True(EventRow(report, EventKind.Policy).CountEntries);
        Assert.Equal(1, EventRow(report, EventKind.Policy).Count);
        Assert.Equal(1, EventRow(report, EventKind.Classic).Count);

        var apple = report.Schools.Single(school => school.SchoolName == "Apple Valley");
        Assert.Equal(2, apple.Counts[EventKind.PublicForum]);
        Assert.Equal(2, apple.Total);

        var wayzata = report.Schools.Single(school => school.SchoolName == "Wayzata");
        Assert.Equal(1, wayzata.Counts[EventKind.Policy]);
        Assert.Equal(1, wayzata.Counts[EventKind.Classic]);
        Assert.Equal(2, wayzata.Total);
    }

    [Fact]
    public void Unlinked_names_do_not_double_count_a_linked_student()
    {
        var report = ParticipationAnalytics.Calculate(new[]
        {
            Entry(EventKind.LincolnDouglas, 1, 10, "Apple Valley", (5, "jane doe")),
            Entry(EventKind.LincolnDouglas, 2, 10, "Apple Valley", (null, "Jane Doe")),
            Entry(EventKind.LincolnDouglas, 3, 11, "Edina", (null, "jane doe")),
            Entry(EventKind.LincolnDouglas, 4, 11, "Edina", (null, "jane doe"))
        });

        Assert.Equal(2, EventRow(report, EventKind.LincolnDouglas).Count);
        Assert.Equal(2, EventRow(report, EventKind.LincolnDouglas).SchoolCount);
        Assert.Equal(new[] { "Apple Valley", "Edina" }, report.Schools.Select(school => school.SchoolName).ToArray());
    }

    [Fact]
    public void An_entry_with_no_competitors_still_counts()
    {
        var report = ParticipationAnalytics.Calculate(new[]
        {
            Entry(EventKind.PublicForum, 9, 4, "Minnetonka")
        });

        Assert.Equal(1, EventRow(report, EventKind.PublicForum).Count);
        Assert.Equal(1, report.Schools.Single().Total);
    }

    [Fact]
    public void Every_event_is_listed_when_the_season_is_empty()
    {
        var report = ParticipationAnalytics.Calculate(Array.Empty<ParticipationEntry>());

        Assert.Equal(ParticipationAnalytics.EventOrder, report.Events.Select(row => row.Event).ToArray());
        Assert.All(report.Events, row => Assert.Equal(0, row.Count));
        Assert.Empty(report.Schools);
        Assert.Equal(0, report.StudentCount);
    }

    [Fact]
    public void Statewide_student_count_keeps_a_student_who_enters_two_events()
    {
        var report = ParticipationAnalytics.Calculate(new[]
        {
            Entry(EventKind.LincolnDouglas, 1, 10, "Apple Valley", (5, "jane doe")),
            Entry(EventKind.PublicForum, 2, 10, "Apple Valley", (5, "jane doe"), (6, "alex kim")),
            Entry(EventKind.Congress, 3, 11, "Edina", (7, "sam lee"))
        });

        Assert.Equal(3, report.StudentCount);
        Assert.Equal(2, report.Schools.Count);
        Assert.Equal("2 schools, 3 students", ParticipationAnalytics.SchoolsAndStudents(report.Schools.Count, report.StudentCount));
    }

    private static EventParticipation EventRow(ParticipationReport report, EventKind kind) =>
        report.Events.Single(row => row.Event == kind);

    private static ParticipationEntry Entry(EventKind kind, int entryId, int schoolId, string school, params (int? Id, string Name)[] competitors) =>
        new(
            kind,
            entryId,
            schoolId,
            school,
            competitors.Select(competitor => new ParticipationCompetitor(competitor.Id, competitor.Name)).ToList());
}
