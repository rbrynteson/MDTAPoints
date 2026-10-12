using MDTAPoints.Core;
using MDTAPoints.Tabroom;

namespace MDTAPoints.Tests;

public class TabroomParserTests
{
    [Fact]
    public void Eden_Prairie_tournament_is_an_in_state_invitational()
    {
        var info = TabroomParser.ParseTournament(Read("tournament-41328.json"));

        Assert.Equal(41328, info.Id);
        Assert.Equal("Eden Prairie Eagle Invitational", info.Name);
        Assert.Equal(new DateOnly(2026, 10, 10), info.StartDate);
        Assert.True(info.InState);
        Assert.Equal(TournamentKind.Invitational.ToString(), info.SuggestedKind);
    }

    [Fact]
    public void Event_list_keeps_varsity_and_open_debate_events()
    {
        var events = TabroomParser.ParseEvents(Read("events-41328.json"));
        var imported = events.Where(item => item.ShouldImport).Select(item => item.Name).OrderBy(name => name).ToArray();

        Assert.Equal(8, events.Count);
        Assert.Equal(new[] { "SENATE Congress (Varsity)", "Varsity Lincoln-Douglas", "Varsity Public Forum" }, imported);
        Assert.Equal(464773, events.Single(item => item.Name.Contains("SENATE")).ResultSetId);
        Assert.Equal(464778, events.Single(item => item.Name.Contains("Lincoln-Douglas") && item.ShouldImport).ResultSetId);
    }

    [Fact]
    public void A_policy_event_without_final_places_uses_the_prelim_seed_list()
    {
        const string eventsJson = """
            {
              "1": {
                "id": 385602,
                "name": "Varsity",
                "level": "open",
                "type": "debate",
                "nsdacategory": 103,
                "ResultSets": [
                  {"id": 463193, "tag": "bracket", "label": "Bracket"},
                  {"id": 463186, "tag": "seed", "label": "Prelim Seeds"},
                  {"id": 463187, "tag": "speaker", "label": "Speaker Awards"},
                  {"id": 463188, "tag": "table", "label": "Prelims Table"}
                ]
              }
            }
            """;
        const string seedJson = """
            [{
              "id": 463186,
              "tag": "seed",
              "label": "Prelim Seeds",
              "Event": {"id": 385602, "name": "Varsity", "type": "debate", "nsdacategory": 103},
              "headers": {"1": {"tag": "WinPm", "description": "Wins in Prelims", "Protocol": {"name": "Prelims"}}},
              "results": [
                {"rank": 1, "place": "1", "Entry": {"id": 1, "name": "Mulrooney & Berger"}, "School": {"id": 10, "name": "St. Paul Central HS"}},
                {"rank": 2, "place": "2", "Entry": {"id": 2, "name": "Washington & Eichenberger"}, "School": {"id": 10, "name": "St. Paul Central HS"}}
              ]
            }]
            """;

        var summary = TabroomParser.ParseEvents(eventsJson).Single();
        var parsed = TabroomParser.ParseResultSet(seedJson);

        Assert.True(summary.ShouldImport);
        Assert.Equal(463186, summary.ResultSetId);
        Assert.Equal(2, parsed.EntryCount);
        Assert.Equal(new[] { "Mulrooney", "Berger" }, parsed.Entries[0].Competitors.Select(competitor => competitor.Name).ToArray());
        Assert.All(parsed.Entries, entry => Assert.False(entry.ReachedFinals));
        Assert.All(parsed.Entries, entry => Assert.False(entry.ReachedElims));
    }

    [Fact]
    public void Senate_results_distinguish_finalists_from_prelims()
    {
        var parsed = TabroomParser.ParseResultSet(Read("senate-464773.json"));

        Assert.Equal(EventKind.Congress.ToString(), parsed.EventKind);
        Assert.Equal(33, parsed.EntryCount);
        Assert.True(parsed.SchoolCount >= 4);
        Assert.True(parsed.HasFinalSession);
        Assert.False(parsed.HasSemis);
        Assert.False(parsed.IsHouse);
        Assert.Equal(11, parsed.Entries.Count(entry => entry.ReachedFinals));
        Assert.Equal(4, Score(parsed, 1).Points);
        Assert.Equal(3, Score(parsed, 6).Points);
        Assert.Equal(2, Score(parsed, 11).Points);
        Assert.Equal(1, Score(parsed, 12).Points);
        Assert.False(parsed.Entries.Single(entry => entry.Place == 12).ReachedFinals);
    }

    [Fact]
    public void Varsity_ld_without_a_break_uses_placement_points()
    {
        var parsed = TabroomParser.ParseResultSet(Read("vld-464778.json"));

        Assert.Equal(EventKind.LincolnDouglas.ToString(), parsed.EventKind);
        Assert.Equal(40, parsed.EntryCount);
        Assert.False(parsed.BreaksToOctosOrDeeper);
        Assert.False(parsed.HasFinalSession);
        Assert.All(parsed.Entries, entry => Assert.True(entry.IsPrelim));
        Assert.Equal(4, Score(parsed, 1).Points);
        Assert.Equal(3, Score(parsed, 6).Points);
        Assert.Equal(2, Score(parsed, 12).Points);
        Assert.Equal(1, Score(parsed, 13).Points);
    }

    private static ScoreResult Score(ImportedEvent parsed, int place)
    {
        var entry = parsed.Entries.Single(item => item.Place == place);
        return CupScorer.Score(new ScoreRequest(
            TournamentKind.Invitational,
            Enum.Parse<EventKind>(parsed.EventKind),
            parsed.CountedDivision,
            true,
            parsed.SchoolCount,
            parsed.EntryCount,
            parsed.BreaksToOctosOrDeeper,
            parsed.HasFinalSession,
            parsed.HasSemis,
            entry.Place,
            entry.ReachedElims,
            entry.ElimWins,
            entry.ReachedFinals,
            entry.ReachedSemis,
            parsed.IsHouse));
    }

    private static string Read(string name) =>
        File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "Fixtures", name));
}
