using MDTAPoints.Api.Data;
using MDTAPoints.Api.Pages.Admin.Tournaments;

namespace MDTAPoints.Tests;

public class ReviewDivisionTests
{
    [Fact]
    public void Events_are_grouped_so_varsity_comes_before_jv_and_novice()
    {
        var groups = ReviewModel.GroupDivisions(new[]
        {
            Event("JV Public Forum", "PublicForum", 1),
            Event("Varsity Lincoln-Douglas", "LincolnDouglas", 2),
            Event("Novice Lincoln-Douglas", "LincolnDouglas", 3),
            Event("Senate (JV-Varsity)", "Congress", 4)
        });

        Assert.Equal(new[] { "Varsity", "JV", "Novice" }, groups.Select(group => group.Label).ToArray());
        Assert.Equal(new[] { "Varsity Lincoln-Douglas", "Senate (JV-Varsity)" }, groups[0].Events.Select(item => item.Name).ToArray());
        Assert.Equal("JV Public Forum", groups[1].Events[0].Name);
        Assert.Equal("Novice Lincoln-Douglas", groups[2].Events[0].Name);
    }

    [Fact]
    public void A_possible_name_is_counted_on_its_event()
    {
        var item = Event("Varsity Policy", "Policy", 1);
        item.Entries.Add(new EntryRow
        {
            Competitors =
            {
                new CompetitorRow { LinkStatus = "Possible" },
                new CompetitorRow { LinkStatus = "Linked" }
            }
        });

        Assert.Equal(1, ReviewModel.NamesToFix(item));
    }

    private static EventRow Event(string name, string kind, int id) => new()
    {
        Id = id,
        Name = name,
        EventKind = kind
    };
}
