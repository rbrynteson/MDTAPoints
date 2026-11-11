using MDTAPoints.Core;

namespace MDTAPoints.Tests;

public class EventClassifierTests
{
    [Theory]
    [InlineData("Senate (JV-Varsity)", "open", true)]
    [InlineData("JV/Varsity Public Forum", null, true)]
    [InlineData("SENATE Congress (Varsity)", "open", true)]
    [InlineData("Varsity Lincoln-Douglas", "open", true)]
    [InlineData("JV Lincoln-Douglas", "open", false)]
    [InlineData("Novice Lincoln-Douglas", "open", false)]
    [InlineData("HOUSE Congress (Novice JV)", "novice", false)]
    [InlineData("Junior Varsity Public Forum", null, false)]
    [InlineData("JV Public Forum", "jv", false)]
    [InlineData("Lincoln-Douglas", "jv", false)]
    [InlineData("Lincoln-Douglas", "open", true)]
    public void Division_guess_treats_combined_jv_varsity_as_the_upper_event(string name, string? level, bool counted)
    {
        Assert.Equal(counted, EventClassifier.IsCountedDivision(name, level));
    }

    [Theory]
    [InlineData("Novice Lincoln-Douglas", null, Division.Novice)]
    [InlineData("JV Public Forum", "open", Division.JuniorVarsity)]
    [InlineData("Junior Varsity Public Forum", null, Division.JuniorVarsity)]
    [InlineData("Senate (JV-Varsity)", "open", Division.Varsity)]
    [InlineData("Lincoln-Douglas", "jv", Division.JuniorVarsity)]
    [InlineData("Lincoln-Douglas", "novice", Division.Novice)]
    [InlineData("Varsity Policy", null, Division.Varsity)]
    public void Division_is_varsity_jv_or_novice(string name, string? level, Division division)
    {
        Assert.Equal(division, EventClassifier.ClassifyDivision(name, level));
        Assert.Equal(division == Division.Varsity, EventClassifier.IsCountedDivision(name, level));
    }

    [Fact]
    public void Senate_jv_varsity_is_a_congress_event_worth_importing()
    {
        Assert.Equal(EventKind.Congress, EventClassifier.Classify("Senate (JV-Varsity)", 301, "congress"));
        Assert.True(EventClassifier.ShouldImport("Senate (JV-Varsity)", 301, "congress", "open"));
        Assert.False(EventClassifier.ShouldImport("JV Lincoln-Douglas", 102, "debate", "open"));
    }
}
