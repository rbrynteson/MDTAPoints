using MDTAPoints.Api.Data;
using MDTAPoints.Api.Services;
using MDTAPoints.Core;

namespace MDTAPoints.Tests;

public class TournamentResultsTests
{
    [Fact]
    public void Results_follow_place_order_and_the_cup_scale_for_that_field()
    {
        var tournament = Tournament();
        var item = tournament.Events[0];

        var rows = TournamentResults.Rows(tournament, item);

        Assert.Equal(new[] { "Ada Lovelace", "Grace Hopper & Jordan Lee", "Prelim" }, rows.Select(row => row.Name).ToArray());
        Assert.Equal(new[] { 4, 3, 1 }, rows.Select(row => row.Points).ToArray());
        Assert.Equal("Hidden semis", rows[0].PlaceNote);
        Assert.Contains("Placed first", rows[0].Scoring);
        Assert.Contains("Finished in the top 4", rows[1].Scoring);
        Assert.Contains("fewer than 20", TournamentResults.Rules(tournament, item)[0]);
    }

    [Fact]
    public void Elimination_wins_score_when_the_event_broke_to_octofinals()
    {
        var item = new EventRow
        {
            Name = "Varsity LD",
            EventKind = nameof(EventKind.LincolnDouglas),
            CountedDivision = true,
            SchoolCount = 8,
            EntryCount = 16,
            BreaksToOctosOrDeeper = true,
            Entries = new List<EntryRow>
            {
                new()
                {
                    Place = 2,
                    SchoolName = "Wayzata",
                    ReachedElims = true,
                    ElimWins = 1,
                    Competitors = new List<CompetitorRow> { new() { RawName = "Grace Hopper" } }
                }
            }
        };
        var tournament = new TournamentDetail
        {
            Kind = nameof(TournamentKind.Invitational),
            InState = true,
            Events = new List<EventRow> { item }
        };

        var row = Assert.Single(TournamentResults.Rows(tournament, item));

        Assert.Equal(3, row.Points);
        Assert.Contains("1 elimination win", row.Scoring);
        Assert.Contains("octofinals", TournamentResults.Rules(tournament, item)[0]);
    }

    [Fact]
    public void Events_follow_the_standings_order()
    {
        var events = TournamentResults.Order(new[]
        {
            Event("Congressional Debate", nameof(EventKind.Congress), 4),
            Event("Policy", nameof(EventKind.Policy), 3),
            Event("Classic", nameof(EventKind.Classic), 2),
            Event("LD", nameof(EventKind.LincolnDouglas), 1)
        });

        Assert.Equal(new[] { "LD", "Classic", "Policy", "Congressional Debate" }, events.Select(item => item.Name).ToArray());
    }

    private static TournamentDetail Tournament()
    {
        return new TournamentDetail
        {
            Name = "Edina",
            Kind = nameof(TournamentKind.Invitational),
            InState = true,
            Events = new List<EventRow> { Event("Varsity LD", nameof(EventKind.LincolnDouglas), 1) }
        };
    }

    private static EventRow Event(string name, string kind, int id)
    {
        return new EventRow
        {
            Id = id,
            Name = name,
            EventKind = kind,
            CountedDivision = true,
            SchoolCount = 8,
            EntryCount = 16,
            Entries = new List<EntryRow>
            {
                new()
                {
                    Place = 9,
                    EntryName = "Prelim",
                    SchoolName = "Minnetonka",
                    Competitors = new List<CompetitorRow> { new() { RawName = "Prelim" } }
                },
                new()
                {
                    Place = 1,
                    EntryName = "Ada",
                    SchoolName = "Edina",
                    PlaceNote = "Hidden semis",
                    Competitors = new List<CompetitorRow> { new() { RawName = "Ada Lovelace" } }
                },
                new()
                {
                    Place = 2,
                    EntryName = "Team",
                    SchoolName = "Wayzata",
                    Competitors = new List<CompetitorRow>
                    {
                        new() { RawName = "Grace Hopper" },
                        new() { RawName = "Jordan Lee" }
                    }
                }
            }
        };
    }
}
