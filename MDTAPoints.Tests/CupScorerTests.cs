using MDTAPoints.Core;

namespace MDTAPoints.Tests;

public class CupScorerTests
{
    [Theory]
    [InlineData(1, 4)]
    [InlineData(6, 3)]
    [InlineData(12, 2)]
    [InlineData(13, 1)]
    [InlineData(40, 1)]
    public void Large_field_without_elims_stacks_place_bonuses(int place, int expected)
    {
        var result = ScoreDebate(place, entryCount: 40, breaksToOctos: false);

        Assert.True(result.Counts);
        Assert.Equal(expected, result.Points);
    }

    [Theory]
    [InlineData(1, 4)]
    [InlineData(4, 3)]
    [InlineData(5, 2)]
    [InlineData(9, 1)]
    public void Field_under_20_uses_top_half_top_four_and_first(int place, int expected)
    {
        var result = ScoreDebate(place, entryCount: 16, breaksToOctos: false);

        Assert.Equal(expected, result.Points);
    }

    [Fact]
    public void Field_under_20_that_breaks_to_octos_uses_elimination_wins()
    {
        var champion = ScoreDebate(1, entryCount: 16, breaksToOctos: true, reachedElims: true, elimWins: 4);
        var breaker = ScoreDebate(9, entryCount: 16, breaksToOctos: true, reachedElims: true, elimWins: 1);
        var prelim = ScoreDebate(10, entryCount: 16, breaksToOctos: true);

        Assert.Equal(5, champion.Points);
        Assert.Contains("Capped at 5 points", champion.Reasons);
        Assert.Equal(3, breaker.Points);
        Assert.Equal(1, prelim.Points);
    }

    [Theory]
    [InlineData(9, 3)]
    [InlineData(4, 2)]
    public void Size_minimums_reject_the_tournament(int schools, int entries)
    {
        var result = CupScorer.Score(Debate() with { SchoolCount = schools, EntryCount = entries, Place = 1 });

        Assert.False(result.Counts);
        Assert.Equal(0, result.Points);
    }

    [Theory]
    [InlineData(1, true, 4)]
    [InlineData(6, true, 3)]
    [InlineData(11, true, 2)]
    [InlineData(12, false, 1)]
    public void Congress_finalists_score_from_the_super_congress(int place, bool reachedFinals, int expected)
    {
        var result = CupScorer.Score(Congress() with { Place = place, ReachedFinals = reachedFinals, EntryCount = 33 });

        Assert.Equal(expected, result.Points);
    }

    [Fact]
    public void Congress_without_a_final_session_is_participation_only()
    {
        var result = CupScorer.Score(Congress() with { HasFinalSession = false, Place = 1, ReachedFinals = false });

        Assert.Equal(1, result.Points);
    }

    [Theory]
    [InlineData(1, 2)]
    [InlineData(4, 2)]
    [InlineData(5, 1)]
    public void Congress_under_12_entries_awards_only_a_top_four_bonus(int place, int expected)
    {
        var result = CupScorer.Score(Congress() with { EntryCount = 8, Place = place, ReachedFinals = place <= 4 });

        Assert.Equal(expected, result.Points);
    }

    [Fact]
    public void Congress_semifinals_award_the_advancing_point_only_in_a_large_field()
    {
        var large = CupScorer.Score(Congress() with
        {
            EntryCount = 90,
            HasSemis = true,
            ReachedFinals = false,
            ReachedSemis = true,
            Place = 20
        });
        var small = CupScorer.Score(Congress() with
        {
            EntryCount = 40,
            HasSemis = true,
            ReachedFinals = false,
            ReachedSemis = true,
            Place = 20
        });

        Assert.Equal(2, large.Points);
        Assert.Equal(1, small.Points);
    }

    [Fact]
    public void Nsda_qualifier_uses_the_shorter_congress_scale()
    {
        var houseWinner = CupScorer.Score(Congress() with
        {
            TournamentKind = TournamentKind.NsdaQualifier,
            IsHouse = true,
            ReachedFinals = true,
            Place = 1
        });
        var senateTopHalf = CupScorer.Score(Congress() with
        {
            TournamentKind = TournamentKind.NsdaQualifier,
            IsHouse = false,
            Place = 5,
            EntryCount = 20
        });
        var policy = CupScorer.Score(Debate() with { TournamentKind = TournamentKind.NsdaQualifier, Place = 1 });

        Assert.Equal(3, houseWinner.Points);
        Assert.Equal(2, senateTopHalf.Points);
        Assert.False(policy.Counts);
    }

    [Theory]
    [InlineData(TournamentKind.Section)]
    [InlineData(TournamentKind.State)]
    public void Section_and_state_results_do_not_score(TournamentKind kind)
    {
        var result = CupScorer.Score(Debate() with { TournamentKind = kind, Place = 1 });

        Assert.False(result.Counts);
    }

    [Fact]
    public void Novice_and_out_of_state_results_do_not_score()
    {
        Assert.False(CupScorer.Score(Debate() with { VarsityOrOpen = false }).Counts);
        Assert.False(CupScorer.Score(Debate() with { InState = false }).Counts);
    }

    private static ScoreResult ScoreDebate(int place, int entryCount, bool breaksToOctos, bool reachedElims = false, int elimWins = 0) =>
        CupScorer.Score(Debate() with
        {
            Place = place,
            EntryCount = entryCount,
            BreaksToOctosOrDeeper = breaksToOctos,
            ReachedElims = reachedElims,
            ElimWins = elimWins
        });

    private static ScoreRequest Debate() => new(
        TournamentKind.Invitational,
        EventKind.LincolnDouglas,
        VarsityOrOpen: true,
        InState: true,
        SchoolCount: 8,
        EntryCount: 40,
        BreaksToOctosOrDeeper: false,
        HasFinalSession: false,
        HasSemis: false,
        Place: 1,
        ReachedElims: false,
        ElimWins: 0,
        ReachedFinals: false,
        ReachedSemis: false,
        IsHouse: false);

    private static ScoreRequest Congress() => Debate() with
    {
        Event = EventKind.Congress,
        EntryCount = 33,
        HasFinalSession = true,
        ReachedFinals = true
    };
}
