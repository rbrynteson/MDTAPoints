namespace MDTAPoints.Core;

public sealed record EventScoringContext(
    TournamentKind TournamentKind,
    EventKind Event,
    bool VarsityOrOpen,
    bool InState,
    int SchoolCount,
    int EntryCount,
    bool BreaksToOctosOrDeeper,
    bool HasFinalSession,
    bool HasSemis,
    bool IsHouse);

public static class ScoringGuide
{
    public static IReadOnlyList<string> Describe(EventScoringContext context)
    {
        if (!context.VarsityOrOpen)
        {
            return new[] { "Novice and JV events do not earn Cup points." };
        }

        if (!context.InState)
        {
            return new[] { "Out-of-state tournaments do not earn Cup points." };
        }

        if (context.TournamentKind == TournamentKind.State)
        {
            return new[] { "MSHSL State results do not earn Cup points." };
        }

        if (context.TournamentKind == TournamentKind.Section)
        {
            return new[] { "Section results do not earn Cup points." };
        }

        if (context.TournamentKind == TournamentKind.NsdaQualifier)
        {
            if (context.Event != EventKind.Congress)
            {
                return new[] { "Only Congressional Debate earns Cup points at an NSDA qualifier." };
            }

            return DescribeNsda(context);
        }

        if (context.SchoolCount < 4)
        {
            return new[] { $"This event had {Schools(context.SchoolCount)}. An event needs at least 4 schools to earn Cup points." };
        }

        if (context.Event == EventKind.Congress)
        {
            return DescribeCongress(context);
        }

        if (context.EntryCount < 10)
        {
            return new[] { $"This event had {Entries(context.EntryCount)}. Debate events need at least 10 entries to earn Cup points." };
        }

        return DescribeDebate(context);
    }

    private static IReadOnlyList<string> DescribeDebate(EventScoringContext context)
    {
        if (context.BreaksToOctosOrDeeper)
        {
            return new[]
            {
                $"This event had {Entries(context.EntryCount)} and broke to octofinals or deeper, so elimination rounds replace place bonuses.",
                "Everyone who competed earns 1 point.",
                "Reaching elimination rounds earns 1 more point.",
                "Each elimination win earns 1 more point.",
                $"Points are capped at {CupScorer.PointCap}."
            };
        }

        if (context.EntryCount >= 20)
        {
            return new[]
            {
                $"This event had {Entries(context.EntryCount)}, so places score on the large-field scale.",
                "Everyone who competed earns 1 point.",
                "Finishing 1st through 12th earns 1 more point.",
                "Finishing 1st through 6th earns 1 more point.",
                "Placing first earns 1 more point."
            };
        }

        return new[]
        {
            $"This event had {Entries(context.EntryCount)}, fewer than 20, so places score on the small-field scale.",
            "Everyone who competed earns 1 point.",
            TopHalf(context.EntryCount),
            "Finishing in the top 4 earns 1 more point.",
            "Placing first earns 1 more point."
        };
    }

    private static IReadOnlyList<string> DescribeCongress(EventScoringContext context)
    {
        if (context.EntryCount < 12)
        {
            return new[]
            {
                $"This event had {Entries(context.EntryCount)}, so Congressional Debate uses the small-field scale.",
                "Everyone who competed earns 1 point.",
                "Finishing in the top 4 earns 1 more point."
            };
        }

        if (!context.HasFinalSession)
        {
            return new[]
            {
                $"This event had {Entries(context.EntryCount)} and no finals.",
                "Everyone who competed earns 1 point.",
                "There is no placement bonus without finals."
            };
        }

        var lines = new List<string>();
        if (context.EntryCount > 80 && context.HasSemis)
        {
            lines.Add($"This event had {Entries(context.EntryCount)}, with semifinals and finals.");
            lines.Add("Everyone who competed earns 1 point.");
            lines.Add("Advancing to semifinals or finals earns 1 more point.");
        }
        else
        {
            lines.Add($"This event had {Entries(context.EntryCount)} and finals.");
            lines.Add("Everyone who competed earns 1 point.");
            lines.Add("Advancing to finals earns 1 more point.");
        }

        lines.Add("Finishing 1st through 6th in finals earns 1 more point.");
        lines.Add("Placing first in finals earns 1 more point.");
        return lines;
    }

    private static IReadOnlyList<string> DescribeNsda(EventScoringContext context)
    {
        if (context.IsHouse)
        {
            return new[]
            {
                "This is an NSDA qualifier House.",
                "Everyone who attended earns 1 point.",
                "Reaching the House final earns 1 more point.",
                "Placing first earns 1 more point."
            };
        }

        return new[]
        {
            $"This is an NSDA qualifier Senate with {Entries(context.EntryCount)}.",
            "Everyone who attended earns 1 point.",
            TopHalf(context.EntryCount),
            "Placing first earns 1 more point."
        };
    }

    private static string TopHalf(int entryCount)
    {
        if (entryCount < 1)
        {
            return "Finishing in the top half earns 1 more point.";
        }

        var cutoff = (int)Math.Ceiling(entryCount / 2.0);
        return $"Finishing in the top half, 1st through {cutoff}, earns 1 more point.";
    }

    private static string Entries(int count) => count == 1 ? "1 entry" : $"{count} entries";

    private static string Schools(int count) => count == 1 ? "1 school" : $"{count} schools";
}
