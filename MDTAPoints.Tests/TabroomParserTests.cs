using System.Reflection;
using MDTAPoints.Api.Pages.Admin.Tournaments;
using MDTAPoints.Api.Services;
using MDTAPoints.Core;
using MDTAPoints.Tabroom;
using Microsoft.AspNetCore.Mvc;

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
    public void Combined_jv_varsity_senate_is_selected_for_import()
    {
        const string eventsJson = """
            {
              "1": {
                "id": 387196,
                "name": "Senate (JV-Varsity)",
                "level": "open",
                "type": "congress",
                "nsdacategory": 301,
                "ResultSets": [{"id": 468748, "tag": "final", "label": "Final Places"}]
              },
              "2": {
                "id": 388202,
                "name": "JV Lincoln Douglas",
                "level": "jv",
                "type": "debate",
                "nsdacategory": 102,
                "ResultSets": [{"id": 468753, "tag": "final", "label": "Final Places"}]
              }
            }
            """;
        const string senateJson = """
            [{
              "id": 468748,
              "tag": "final",
              "label": "Final Places",
              "Event": {"id": 387196, "name": "Senate (JV-Varsity)", "level": "open", "type": "congress", "nsdacategory": 301},
              "headers": {"1": {"tag": "Place", "description": "Final Place", "Protocol": {"name": "Finals"}}},
              "results": [
                {"rank": 1, "place": "1", "Entry": {"id": 1, "name": "Alex Kim"}, "School": {"id": 10, "name": "Minnetonka"}}
              ]
            }]
            """;

        var events = TabroomParser.ParseEvents(eventsJson);
        var senate = events.Single(item => item.Name.Contains("Senate"));
        var parsed = TabroomParser.ParseResultSet(senateJson);

        Assert.True(senate.ShouldImport);
        Assert.Equal(468748, senate.ResultSetId);
        Assert.False(events.Single(item => item.Name.StartsWith("JV")).ShouldImport);
        Assert.True(parsed.CountedDivision);
        Assert.Equal(Division.Varsity.ToString(), parsed.Division);
        Assert.Equal(EventKind.Congress.ToString(), parsed.EventKind);

        var rows = ImportModel.BuildRows(events);
        Assert.Equal("Varsity", rows.Single(row => row.Name.Contains("Senate")).Cup);
        Assert.True(rows.Single(row => row.Name.Contains("Senate")).Counts);
        Assert.True(rows.Single(row => row.Name.Contains("Senate")).Suggested);
        Assert.Equal("Novice or JV, not in the Cup", rows.Single(row => row.Name.StartsWith("JV")).Cup);
        Assert.False(rows.Single(row => row.Name.StartsWith("JV")).Counts);
        Assert.True(rows.Single(row => row.Name.StartsWith("JV")).Suggested);

        var house = TabroomImportService.ApplyVarsityChoice(parsed, false);
        Assert.False(house.CountedDivision);
        Assert.Equal(Division.Varsity.ToString(), house.Division);
        var countedJv = TabroomImportService.ApplyVarsityChoice(
            TabroomParser.ParseResultSet(senateJson.Replace("Senate (JV-Varsity)", "JV Lincoln Douglas")),
            true);
        Assert.True(countedJv.CountedDivision);
        Assert.Equal(Division.Varsity.ToString(), countedJv.Division);
    }

    [Fact]
    public void Varsity_choices_bind_under_their_own_form_name()
    {
        var name = typeof(ImportModel).GetProperty(nameof(ImportModel.Varsity))!
            .GetCustomAttribute<BindPropertyAttribute>()!
            .Name;

        Assert.Equal("Varsity", name);
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

    [Fact]
    public void A_break_to_triples_scores_elimination_wins_instead_of_placement()
    {
        const string json = """
            [{
              "id": 474077,
              "tag": "final",
              "label": "Final Places",
              "Event": {"id": 360612, "name": "Varsity Lincoln-Douglas", "level": "open", "type": "debate", "nsdacategory": 102},
              "headers": {"1": {"tag": "Round", "description": "Last round competed in", "Protocol": {"name": "Debate Elims"}}},
              "results": [
                {"rank": 1, "place": "1st", "Entry": {"id": 1, "name": "Ada"}, "School": {"id": 1, "name": "Edina"}, "values": {"1": "Finals"}},
                {"rank": 2, "place": "2nd", "Entry": {"id": 2, "name": "Grace"}, "School": {"id": 2, "name": "Eagan"}, "values": {"1": "Finals"}},
                {"rank": 3, "place": "Semis", "Entry": {"id": 3, "name": "Lin"}, "School": {"id": 3, "name": "Wayzata"}, "values": {"1": "Semis"}},
                {"rank": 4, "place": "Quarters", "Entry": {"id": 4, "name": "Maya"}, "School": {"id": 4, "name": "Minnetonka"}, "values": {"1": "Quarters"}},
                {"rank": 5, "place": "Octos", "Entry": {"id": 5, "name": "Noor"}, "School": {"id": 1, "name": "Edina"}, "values": {"1": "Octos"}},
                {"rank": 6, "place": "Doubles", "Entry": {"id": 6, "name": "Owen"}, "School": {"id": 2, "name": "Eagan"}, "values": {"1": "Doubles"}},
                {"rank": 7, "place": "Triples", "Entry": {"id": 7, "name": "Priya"}, "School": {"id": 3, "name": "Wayzata"}, "values": {"1": "Triples"}},
                {"rank": 8, "place": "Prelim", "Entry": {"id": 8, "name": "Quinn"}, "School": {"id": 4, "name": "Minnetonka"}, "values": {"1": "Prelim"}},
                {"rank": 9, "place": "Prelim", "Entry": {"id": 9, "name": "Rae"}, "School": {"id": 1, "name": "Edina"}},
                {"rank": 10, "place": "Prelim", "Entry": {"id": 10, "name": "Sam"}, "School": {"id": 2, "name": "Eagan"}}
              ]
            }]
            """;

        var parsed = TabroomParser.ParseResultSet(json);

        Assert.True(parsed.BreaksToOctosOrDeeper);
        Assert.True(parsed.HasSemis);
        Assert.False(parsed.HasFinalSession);
        Assert.Equal(6, Wins(parsed, "Ada"));
        Assert.Equal(5, Wins(parsed, "Grace"));
        Assert.Equal(4, Wins(parsed, "Lin"));
        Assert.Equal(3, Wins(parsed, "Maya"));
        Assert.Equal(2, Wins(parsed, "Noor"));
        Assert.Equal(1, Wins(parsed, "Owen"));
        Assert.Equal(0, Wins(parsed, "Priya"));
        Assert.True(parsed.Entries.Single(entry => entry.EntryName == "Priya").ReachedElims);
        Assert.False(parsed.Entries.Single(entry => entry.EntryName == "Quinn").ReachedElims);
        Assert.Equal(new[] { "Ada", "Grace" }, parsed.Entries.Where(entry => entry.ReachedFinals).Select(entry => entry.EntryName).OrderBy(name => name).ToArray());
        Assert.Equal(5, Score(parsed, 1).Points);
        Assert.Equal(5, Score(parsed, 4).Points);
        Assert.Equal(4, Score(parsed, 5).Points);
        Assert.Equal(3, Score(parsed, 6).Points);
        Assert.Equal(2, Score(parsed, 7).Points);
        Assert.Equal(1, Score(parsed, 8).Points);
    }

    [Fact]
    public void Congress_chamber_sheets_keep_finals_separate_from_semis_and_prelims()
    {
        const string eventsJson = """
            {
              "1": {
                "id": 386471,
                "name": "Congressional Debate",
                "level": "open",
                "type": "congress",
                "nsdacategory": 301,
                "ResultSets": [
                  {"id": 11, "tag": "chamber", "label": "Rd Chamber Results"},
                  {"id": 12, "tag": "chamber", "label": "Semi Chamber Results"},
                  {"id": 13, "tag": "chamber", "label": "Final Chamber Results"},
                  {"id": 14, "tag": "qualifier", "label": "TOC Qualifying Bids"}
                ]
              }
            }
            """;
        const string prelimJson = """
            [{"id": 11, "tag": "chamber", "label": "Rd Chamber Results",
              "Event": {"id": 386471, "name": "Congressional Debate", "level": "open", "type": "congress", "nsdacategory": 301},
              "results": [
                {"rank": 1, "place": "1", "Entry": {"id": 1, "name": "Finalist"}, "School": {"id": 1, "name": "Edina"}},
                {"rank": 1, "place": "1", "Entry": {"id": 2, "name": "Semifinalist"}, "School": {"id": 2, "name": "Eagan"}},
                {"rank": 2, "place": "2", "Entry": {"id": 3, "name": "Prelim One"}, "School": {"id": 3, "name": "Wayzata"}},
                {"rank": 2, "place": "2", "Entry": {"id": 4, "name": "Prelim Two"}, "School": {"id": 4, "name": "Minnetonka"}}
              ]}]
            """;
        const string semiJson = """
            [{"id": 12, "tag": "chamber", "label": "Semi Chamber Results",
              "Event": {"id": 386471, "name": "Congressional Debate", "level": "open", "type": "congress", "nsdacategory": 301},
              "results": [
                {"rank": 1, "place": "1", "Entry": {"id": 1, "name": "Finalist"}, "School": {"id": 1, "name": "Edina"}},
                {"rank": 9, "place": "9", "Entry": {"id": 2, "name": "Semifinalist"}, "School": {"id": 2, "name": "Eagan"}}
              ]}]
            """;
        const string finalJson = """
            [{"id": 13, "tag": "chamber", "label": "Final Chamber Results",
              "Event": {"id": 386471, "name": "Congressional Debate", "level": "open", "type": "congress", "nsdacategory": 301},
              "results": [
                {"rank": 1, "place": "1", "Entry": {"id": 1, "name": "Finalist"}, "School": {"id": 1, "name": "Edina"}}
              ]}]
            """;

        var summary = TabroomParser.ParseEvents(eventsJson).Single();
        var parsed = TabroomParser.ParseResultSets(new[] { prelimJson, semiJson, finalJson });

        Assert.Equal(new[] { 11, 12, 13 }, summary.ResultSetIds);
        Assert.Equal(13, summary.ResultSetId);
        Assert.Equal(4, parsed.EntryCount);
        Assert.True(parsed.HasFinalSession);
        Assert.True(parsed.HasSemis);
        Assert.Equal(1, parsed.Entries.Single(entry => entry.EntryName == "Finalist").Place);
        Assert.True(parsed.Entries.Single(entry => entry.EntryName == "Finalist").ReachedFinals);
        Assert.True(parsed.Entries.Single(entry => entry.EntryName == "Finalist").ReachedSemis);
        Assert.False(parsed.Entries.Single(entry => entry.EntryName == "Semifinalist").ReachedFinals);
        Assert.True(parsed.Entries.Single(entry => entry.EntryName == "Semifinalist").ReachedSemis);
        Assert.True(parsed.Entries.Single(entry => entry.EntryName == "Prelim One").IsPrelim);
        Assert.False(parsed.Entries.Single(entry => entry.EntryName == "Prelim One").ReachedSemis);

        var largeFinal = CupScorer.Score(new ScoreRequest(
            TournamentKind.Invitational, EventKind.Congress, true, true, parsed.SchoolCount, 97,
            false, true, true, 1, true, 0, true, true, false));
        var largeSemi = CupScorer.Score(new ScoreRequest(
            TournamentKind.Invitational, EventKind.Congress, true, true, parsed.SchoolCount, 97,
            false, true, true, 2, true, 0, false, true, false));
        var largePrelim = CupScorer.Score(new ScoreRequest(
            TournamentKind.Invitational, EventKind.Congress, true, true, parsed.SchoolCount, 97,
            false, true, true, 3, false, 0, false, false, false));
        Assert.Equal(4, largeFinal.Points);
        Assert.Equal(2, largeSemi.Points);
        Assert.Equal(1, largePrelim.Points);
    }

    [Fact]
    public void A_school_from_outside_Minnesota_is_not_imported_when_its_state_is_posted()
    {
        const string json = """
            [{
              "id": 1,
              "tag": "final",
              "Event": {"id": 9, "name": "Varsity Lincoln-Douglas", "level": "open", "type": "debate", "nsdacategory": 102},
              "results": [
                {"rank": 1, "place": "1", "Entry": {"id": 1, "name": "Ada"}, "School": {"id": 1, "name": "Edina", "state": "MN"}},
                {"rank": 2, "place": "2", "Entry": {"id": 2, "name": "Bea"}, "School": {"id": 2, "name": "The Harker School", "state": "CA"}},
                {"rank": 3, "place": "3", "Entry": {"id": 3, "name": "Cara"}, "School": {"id": 3, "name": "Eagan"}}
              ]
            }]
            """;

        var parsed = TabroomParser.ParseResultSet(json);
        var skipped = new List<string>();
        var kept = TabroomImportService.OmitOutOfStateSchools(parsed, skipped, parsed.Name);

        Assert.Equal(3, kept.EntryCount);
        Assert.Equal(3, kept.SchoolCount);
        Assert.Equal(new[] { "Ada", "Cara" }, kept.Entries.Select(entry => entry.EntryName).ToArray());
        Assert.Contains("The Harker School (CA)", skipped[0]);
        Assert.False(TabroomParser.IncludeSchool("IA"));
        Assert.True(TabroomParser.IncludeSchool(null));
        Assert.True(TabroomParser.IncludeSchool("mn"));
    }

    private static int Wins(ImportedEvent parsed, string name) =>
        parsed.Entries.Single(entry => entry.EntryName == name).ElimWins;

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
