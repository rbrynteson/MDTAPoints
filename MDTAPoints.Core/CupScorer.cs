namespace MDTAPoints.Core;

public static class CupScorer
{
    public const int PointCap = 5;

    public static ScoreResult Score(ScoreRequest request)
    {
        var reasons = new List<string>();

        if (!request.VarsityOrOpen)
        {
            return Skip("Novice and JV events do not score");
        }

        if (!request.InState)
        {
            return Skip("Out-of-state tournaments do not score");
        }

        if (request.TournamentKind == TournamentKind.State)
        {
            return Skip("MSHSL State results do not score");
        }

        if (request.TournamentKind == TournamentKind.Section)
        {
            return Skip("Section results do not score");
        }

        if (request.TournamentKind == TournamentKind.NsdaQualifier)
        {
            if (request.Event != EventKind.Congress)
            {
                return Skip("The NSDA qualifier exception applies to Congressional Debate");
            }

            return Finish(ScoreNsda(request, reasons), reasons);
        }

        if (request.SchoolCount < 4)
        {
            return Skip("Fewer than 4 schools");
        }

        if (request.Event == EventKind.Congress)
        {
            return Finish(ScoreCongress(request, reasons), reasons);
        }

        if (request.EntryCount < 10)
        {
            return Skip("Fewer than 10 entries");
        }

        reasons.Add("Competed");
        var points = 1;

        if (request.BreaksToOctosOrDeeper)
        {
            var reached = request.ReachedElims || request.ElimWins > 0;
            if (reached)
            {
                points++;
                reasons.Add("Reached elimination rounds");
            }

            var wins = Math.Max(0, request.ElimWins);
            if (wins > 0)
            {
                points += wins;
                reasons.Add(wins == 1 ? "1 elimination win" : $"{wins} elimination wins");
            }
        }
        else if (request.EntryCount >= 20)
        {
            if (request.Place is >= 1 and <= 12)
            {
                points++;
                reasons.Add("Finished 1st through 12th");
            }

            if (request.Place is >= 1 and <= 6)
            {
                points++;
                reasons.Add("Finished 1st through 6th");
            }

            if (request.Place == 1)
            {
                points++;
                reasons.Add("Placed first");
            }
        }
        else
        {
            var topHalf = (int)Math.Ceiling(request.EntryCount / 2.0);
            if (request.Place >= 1 && request.Place <= topHalf)
            {
                points++;
                reasons.Add("Finished in the top half");
            }

            if (request.Place is >= 1 and <= 4)
            {
                points++;
                reasons.Add("Finished in the top 4");
            }

            if (request.Place == 1)
            {
                points++;
                reasons.Add("Placed first");
            }
        }

        return Finish(points, reasons);
    }

    private static int ScoreCongress(ScoreRequest request, List<string> reasons)
    {
        reasons.Add("Competed");
        var points = 1;

        if (request.EntryCount < 12)
        {
            if (request.Place is >= 1 and <= 4)
            {
                points++;
                reasons.Add("Finished in the top 4 of a field under 12");
            }

            return points;
        }

        if (!request.HasFinalSession)
        {
            reasons.Add("No Super Congress");
            return points;
        }

        var advanced = request.ReachedFinals || (request.EntryCount > 80 && request.HasSemis && request.ReachedSemis);
        if (advanced)
        {
            points++;
            reasons.Add(request.ReachedFinals
                ? "Advanced to Super Congress"
                : "Advanced to semifinals");
        }

        if (request.ReachedFinals && request.Place is >= 1 and <= 6)
        {
            points++;
            reasons.Add("Finished 1st through 6th in Super Congress");
        }

        if (request.ReachedFinals && request.Place == 1)
        {
            points++;
            reasons.Add("Placed first in Super Congress");
        }

        return points;
    }

    private static int ScoreNsda(ScoreRequest request, List<string> reasons)
    {
        reasons.Add("Attended an NSDA qualifier");
        var points = 1;

        if (request.IsHouse)
        {
            if (request.ReachedFinals)
            {
                points++;
                reasons.Add("Reached the House final");
            }
        }
        else
        {
            var topHalf = (int)Math.Ceiling(request.EntryCount / 2.0);
            if (request.Place >= 1 && request.Place <= topHalf)
            {
                points++;
                reasons.Add("Finished in the top half of Senate");
            }
        }

        if (request.Place == 1)
        {
            points++;
            reasons.Add("Placed first");
        }

        return points;
    }

    private static ScoreResult Finish(int points, List<string> reasons)
    {
        if (points > PointCap)
        {
            points = PointCap;
            reasons.Add("Capped at 5 points");
        }

        return new ScoreResult(points, true, reasons);
    }

    private static ScoreResult Skip(string reason) =>
        new(0, false, new[] { reason });
}
