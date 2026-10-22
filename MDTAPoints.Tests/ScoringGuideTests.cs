using MDTAPoints.Core;

namespace MDTAPoints.Tests;

public class ScoringGuideTests
{
    [Fact]
    public void Debate_under_20_describes_the_top_half_scale()
    {
        var lines = ScoringGuide.Describe(Debate(16));

        Assert.Equal(new[]
        {
            "This event had 16 entries, fewer than 20, so places score on the small-field scale.",
            "Everyone who competed earns 1 point.",
            "Finishing in the top half, 1st through 8, earns 1 more point.",
            "Finishing in the top 4 earns 1 more point.",
            "Placing first earns 1 more point."
        }, lines);
    }

    [Fact]
    public void Debate_with_20_or_more_describes_the_large_field_scale()
    {
        var lines = ScoringGuide.Describe(Debate(20));

        Assert.Equal(new[]
        {
            "This event had 20 entries, so places score on the large-field scale.",
            "Everyone who competed earns 1 point.",
            "Finishing 1st through 12th earns 1 more point.",
            "Finishing 1st through 6th earns 1 more point.",
            "Placing first earns 1 more point."
        }, lines);
    }

    [Fact]
    public void Nineteen_entries_still_use_the_top_half_cutoff()
    {
        var lines = ScoringGuide.Describe(Debate(19));

        Assert.Contains("fewer than 20", lines[0]);
        Assert.Contains("Finishing in the top half, 1st through 10, earns 1 more point.", lines);
    }

    [Fact]
    public void Breaking_to_octofinals_replaces_place_bonuses_and_caps_at_five()
    {
        var lines = ScoringGuide.Describe(Debate(16, octos: true));

        Assert.Equal(new[]
        {
            "This event had 16 entries and broke to octofinals or deeper, so elimination rounds replace place bonuses.",
            "Everyone who competed earns 1 point.",
            "Reaching elimination rounds earns 1 more point.",
            "Each elimination win earns 1 more point.",
            "Points are capped at 5."
        }, lines);
    }

    [Fact]
    public void A_short_field_does_not_score_even_when_it_breaks_to_octofinals()
    {
        var lines = ScoringGuide.Describe(Debate(9, octos: true));

        Assert.Equal(new[] { "This event had 9 entries. Debate events need at least 10 entries to earn Cup points." }, lines);
    }

    [Theory]
    [InlineData(1, "This event had 1 school. An event needs at least 4 schools to earn Cup points.")]
    [InlineData(3, "This event had 3 schools. An event needs at least 4 schools to earn Cup points.")]
    public void Fewer_than_four_schools_does_not_score(int schools, string expected)
    {
        var lines = ScoringGuide.Describe(Debate(30, schools: schools));

        Assert.Equal(new[] { expected }, lines);
    }

    [Fact]
    public void Congress_under_12_awards_only_a_top_four_bonus()
    {
        var lines = ScoringGuide.Describe(Congress(8, finals: true));

        Assert.Equal(new[]
        {
            "This event had 8 entries, so Congressional Debate uses the small-field scale.",
            "Everyone who competed earns 1 point.",
            "Finishing in the top 4 earns 1 more point."
        }, lines);
    }

    [Fact]
    public void Congress_without_super_congress_is_participation_only()
    {
        var lines = ScoringGuide.Describe(Congress(33, finals: false));

        Assert.Equal(new[]
        {
            "This event had 33 entries and no finals.",
            "Everyone who competed earns 1 point.",
            "There is no placement bonus without finals."
        }, lines);
    }

    [Fact]
    public void Congress_with_super_congress_describes_finalist_bonuses()
    {
        var lines = ScoringGuide.Describe(Congress(33, finals: true));

        Assert.Equal(new[]
        {
            "This event had 33 entries and finals.",
            "Everyone who competed earns 1 point.",
            "Advancing to finals earns 1 more point.",
            "Finishing 1st through 6th in finals earns 1 more point.",
            "Placing first in finals earns 1 more point."
        }, lines);
    }

    [Fact]
    public void A_huge_congress_field_counts_a_semifinal_advance()
    {
        var lines = ScoringGuide.Describe(Congress(90, finals: true, semis: true));

        Assert.Contains("with semifinals and finals", lines[0]);
        Assert.Contains("Advancing to semifinals or finals earns 1 more point.", lines);
    }

    [Fact]
    public void Nsda_house_and_senate_use_the_qualifier_scale()
    {
        var house = ScoringGuide.Describe(Congress(20, finals: true, kind: TournamentKind.NsdaQualifier, house: true));
        var senate = ScoringGuide.Describe(Congress(20, finals: true, kind: TournamentKind.NsdaQualifier));

        Assert.Equal(new[]
        {
            "This is an NSDA qualifier House.",
            "Everyone who attended earns 1 point.",
            "Reaching the House final earns 1 more point.",
            "Placing first earns 1 more point."
        }, house);
        Assert.Contains("This is an NSDA qualifier Senate with 20 entries.", senate);
        Assert.Contains("Finishing in the top half, 1st through 10, earns 1 more point.", senate);
    }

    [Fact]
    public void An_nsda_debate_event_does_not_use_the_invitational_scale()
    {
        var lines = ScoringGuide.Describe(Debate(40) with { TournamentKind = TournamentKind.NsdaQualifier });

        Assert.Equal(new[] { "Only Congressional Debate earns Cup points at an NSDA qualifier." }, lines);
    }

    [Theory]
    [InlineData(TournamentKind.State, "MSHSL State results do not earn Cup points.")]
    [InlineData(TournamentKind.Section, "Section results do not earn Cup points.")]
    public void State_and_section_results_do_not_score(TournamentKind kind, string expected)
    {
        var lines = ScoringGuide.Describe(Debate(40) with { TournamentKind = kind });

        Assert.Equal(new[] { expected }, lines);
    }

    [Fact]
    public void Out_of_state_and_novice_events_do_not_score()
    {
        Assert.Equal(
            new[] { "Out-of-state tournaments do not earn Cup points." },
            ScoringGuide.Describe(Debate(40) with { InState = false }));
        Assert.Equal(
            new[] { "Novice and JV events do not earn Cup points." },
            ScoringGuide.Describe(Debate(40) with { VarsityOrOpen = false }));
    }

    private static EventScoringContext Debate(int entries, bool octos = false, int schools = 8) =>
        new(TournamentKind.Invitational, EventKind.LincolnDouglas, true, true, schools, entries, octos, false, false, false);

    private static EventScoringContext Congress(
        int entries,
        bool finals,
        bool semis = false,
        TournamentKind kind = TournamentKind.Invitational,
        bool house = false) =>
        new(kind, EventKind.Congress, true, true, 8, entries, false, finals, semis, house);
}
