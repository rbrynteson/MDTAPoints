using MDTAPoints.Core;

namespace MDTAPoints.Tests;

public class NameLinkerTests
{
    private static readonly KnownStudent[] School =
    {
        new("1", "edina", "matthew bartley", Array.Empty<string>()),
        new("2", "edina", "kate hu", new[] { "katherine hu" })
    };

    [Fact]
    public void Exact_name_and_saved_alias_link_automatically()
    {
        var exact = NameLinker.Classify("Matthew Bartley", "edina", School);
        var alias = NameLinker.Classify("Katherine Hu", "edina", School);

        Assert.Equal(LinkStatus.Linked, exact.Status);
        Assert.Equal("1", exact.StudentId);
        Assert.Equal(LinkStatus.Linked, alias.Status);
        Assert.Equal("2", alias.StudentId);
    }

    [Fact]
    public void Close_names_stay_unresolved_until_someone_confirms_them()
    {
        var decision = NameLinker.Classify("Matt Bartley", "edina", School);

        Assert.Equal(LinkStatus.Possible, decision.Status);
        Assert.Null(decision.StudentId);
        Assert.Equal("1", decision.SuggestedStudentId);
        Assert.False(NameLinker.CanPublish(new[] { decision.Status }));
    }

    [Fact]
    public void A_different_school_or_a_different_student_is_new()
    {
        var otherSchool = NameLinker.Classify("Matthew Bartley", "minnetonka", School);
        var otherStudent = NameLinker.Classify("Sunny Baldwin", "edina", School);

        Assert.Equal(LinkStatus.New, otherSchool.Status);
        Assert.Equal(LinkStatus.New, otherStudent.Status);
        Assert.True(NameLinker.CanPublish(new[] { LinkStatus.Linked, LinkStatus.New }));
    }

    [Fact]
    public void Last_comma_first_normalizes_to_the_same_name()
    {
        Assert.Equal("matthew bartley", NameNormalizer.Normalize("Bartley, Matthew"));
    }
}
