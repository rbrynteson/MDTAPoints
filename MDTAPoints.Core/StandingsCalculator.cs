namespace MDTAPoints.Core;

public static class StandingsCalculator
{
    public const int TournamentsRequired = 4;

    public static SeasonStandings Calculate(SeasonInput input)
    {
        var qualified = input.Qualifications
            .Where(item => item.Qualified)
            .Select(item => (item.StudentId, item.Event))
            .ToHashSet();
        var headToHead = input.HeadToHead
            .Select(item => (item.WinnerStudentId, item.LoserStudentId))
            .ToHashSet();

        var individualApps = new List<Appearance>();
        var teamApps = new List<Appearance>();

        foreach (var entry in input.Entries)
        {
            var score = CupScorer.Score(new ScoreRequest(
                entry.TournamentKind,
                entry.Event,
                entry.VarsityOrOpen,
                entry.InState,
                entry.SchoolCount,
                entry.EntryCount,
                entry.BreaksToOctosOrDeeper,
                entry.HasFinalSession,
                entry.HasSemis,
                entry.Place,
                entry.ReachedElims,
                entry.ElimWins,
                entry.ReachedFinals,
                entry.ReachedSemis,
                entry.IsHouse));

            if (!score.Counts)
            {
                continue;
            }

            var detail = string.Join(", ", score.Reasons);
            if (EventKindLabels.IsTeamEvent(entry.Event))
            {
                teamApps.Add(Appearance.FromEntry(entry, "entry:" + entry.EntryId, score.Points, detail));
            }

            foreach (var competitor in entry.Competitors)
            {
                individualApps.Add(Appearance.FromCompetitor(entry, competitor, score.Points, detail));
                if (!EventKindLabels.IsTeamEvent(entry.Event))
                {
                    teamApps.Add(Appearance.FromCompetitor(entry, competitor, score.Points, detail));
                }
            }
        }

        var individuals = BuildIndividuals(individualApps, qualified, headToHead);
        var teams = BuildTeams(teamApps);
        return new SeasonStandings(individuals, teams);
    }

    private static List<IndividualStanding> BuildIndividuals(
        List<Appearance> apps,
        HashSet<(string StudentId, EventKind Event)> qualified,
        HashSet<(string Winner, string Loser)> headToHead)
    {
        var rows = new List<WorkingIndividual>();
        foreach (var group in apps.GroupBy(app => (app.StudentId, app.Event)))
        {
            var perTournament = group
                .GroupBy(app => app.TournamentId)
                .Select(tournament => tournament.OrderByDescending(app => app.Points).ThenBy(app => app.Place).First())
                .ToList();

            var kept = new List<Appearance>();
            var dropped = new List<Appearance>();
            foreach (var week in perTournament.GroupBy(app => app.Week))
            {
                var ordered = week
                    .OrderByDescending(app => app.Points)
                    .ThenBy(app => app.Place)
                    .ThenBy(app => app.TournamentId, StringComparer.Ordinal)
                    .ToList();
                kept.Add(ordered[0]);
                dropped.AddRange(ordered.Skip(1));
            }

            var latest = (kept.Count > 0 ? kept : perTournament).OrderByDescending(app => app.Start).First();
            var isQualified = qualified.Contains((group.Key.StudentId, group.Key.Event));
            var lines = kept.Concat(dropped)
                .OrderBy(app => app.Start)
                .ThenBy(app => app.TournamentName, StringComparer.OrdinalIgnoreCase)
                .Select(app => app.ToLine(dropped.Contains(app)))
                .ToList();
            if (isQualified)
            {
                lines.Add(new PointLine("", "State qualifier", latest.Start, 1, 0, false, "Qualified to state"));
            }

            rows.Add(new WorkingIndividual
            {
                StudentId = group.Key.StudentId,
                Name = latest.StudentName,
                SchoolId = latest.SchoolId,
                SchoolName = latest.SchoolName,
                Event = group.Key.Event,
                Points = kept.Sum(app => app.Points) + (isQualified ? 1 : 0),
                TournamentCount = kept.Select(app => app.TournamentId).Distinct().Count(),
                QualifiedToState = isQualified,
                Grade = group.Select(app => app.Grade).Where(grade => grade.HasValue).Select(grade => grade!.Value).DefaultIfEmpty().Max(),
                AveragePlace = kept.Count == 0 ? 0 : kept.Average(app => (double)app.Place),
                Lines = lines
            });
        }

        foreach (var row in rows)
        {
            if (row.Grade == 0)
            {
                row.Grade = null;
            }
        }

        var orderedRows = rows
            .OrderByDescending(row => row.Points)
            .ThenByDescending(row => row.QualifiedToState)
            .ThenByDescending(row => row.Grade ?? int.MinValue)
            .ThenBy(row => row.Name, StringComparer.OrdinalIgnoreCase)
            .ToList();

        orderedRows = ApplyIndividualTieBreak(orderedRows, headToHead);

        var awardOrder = orderedRows.Where(row => row.TournamentCount >= TournamentsRequired).ToList();
        for (var index = 0; index < awardOrder.Count; index++)
        {
            var tiedWithPrevious = index > 0 && FullyTied(awardOrder[index - 1], awardOrder[index], headToHead);
            awardOrder[index].AwardRank = tiedWithPrevious ? awardOrder[index - 1].AwardRank : index + 1;
        }

        return orderedRows.Select(row => row.ToStanding()).ToList();
    }

    private static List<WorkingIndividual> ApplyIndividualTieBreak(
        List<WorkingIndividual> rows,
        HashSet<(string Winner, string Loser)> headToHead)
    {
        var result = new List<WorkingIndividual>();
        var index = 0;
        while (index < rows.Count)
        {
            var end = index + 1;
            while (end < rows.Count && SamePrimaryKey(rows[index], rows[end]))
            {
                end++;
            }

            var group = rows.Skip(index).Take(end - index).ToList();
            group.Sort((left, right) => CompareRest(left, right, group.Count, headToHead));
            result.AddRange(group);
            index = end;
        }

        return result;
    }

    private static bool SamePrimaryKey(WorkingIndividual left, WorkingIndividual right) =>
        left.Points == right.Points && left.QualifiedToState == right.QualifiedToState && left.Grade == right.Grade;

    private static int CompareRest(
        WorkingIndividual left,
        WorkingIndividual right,
        int groupSize,
        HashSet<(string Winner, string Loser)> headToHead)
    {
        if (groupSize == 2)
        {
            if (headToHead.Contains((left.StudentId, right.StudentId)))
            {
                return -1;
            }

            if (headToHead.Contains((right.StudentId, left.StudentId)))
            {
                return 1;
            }
        }

        return CompareAward(left, right);
    }

    private static bool FullyTied(
        WorkingIndividual left,
        WorkingIndividual right,
        HashSet<(string Winner, string Loser)> headToHead) =>
        left.Points == right.Points
        && left.QualifiedToState == right.QualifiedToState
        && left.Grade == right.Grade
        && left.TournamentCount == right.TournamentCount
        && Math.Abs(left.AveragePlace - right.AveragePlace) < 0.0001
        && !headToHead.Contains((left.StudentId, right.StudentId))
        && !headToHead.Contains((right.StudentId, left.StudentId));

    private static int CompareAward(WorkingIndividual left, WorkingIndividual right)
    {
        var tournaments = right.TournamentCount.CompareTo(left.TournamentCount);
        if (tournaments != 0)
        {
            return tournaments;
        }

        var place = left.AveragePlace.CompareTo(right.AveragePlace);
        if (place != 0)
        {
            return place;
        }

        return string.Compare(left.Name, right.Name, StringComparison.OrdinalIgnoreCase);
    }

    private static List<TeamStanding> BuildTeams(List<Appearance> apps)
    {
        var rows = new List<WorkingTeam>();
        foreach (var schoolEvent in apps.GroupBy(app => (app.SchoolId, app.Event)))
        {
            var keptLines = new List<PointLine>();
            var droppedLines = new List<PointLine>();
            var places = new List<int>();
            var points = 0;
            var tournaments = 0;

            foreach (var week in schoolEvent.GroupBy(app => app.Week))
            {
                var choices = week.GroupBy(app => app.TournamentId).Select(tournament =>
                {
                    var top = tournament.OrderByDescending(app => app.Points).ThenBy(app => app.Place).Take(2).ToList();
                    return new TournamentChoice(tournament.First(), top, top.Sum(app => app.Points));
                }).OrderByDescending(choice => choice.Sum).ThenBy(choice => choice.Top.Min(app => app.Place)).ToList();

                if (choices.Count == 0)
                {
                    continue;
                }

                points += choices[0].Sum;
                tournaments++;
                places.AddRange(choices[0].Top.Select(app => app.Place));
                keptLines.Add(choices[0].Source.ToLine(false, choices[0].Sum, "Top 2"));
                droppedLines.AddRange(choices.Skip(1).Select(choice => choice.Source.ToLine(true, choice.Sum, "Another tournament this week")));
            }

            var latest = schoolEvent.OrderByDescending(app => app.Start).First();
            rows.Add(new WorkingTeam
            {
                SchoolId = schoolEvent.Key.SchoolId,
                SchoolName = latest.SchoolName,
                Event = schoolEvent.Key.Event,
                Points = points,
                TournamentCount = tournaments,
                AveragePlace = places.Count == 0 ? 0 : places.Average(),
                Lines = keptLines.Concat(droppedLines).OrderBy(line => line.Start).ToList()
            });
        }

        return RankTeams(rows).Select(row => row.ToStanding()).ToList();
    }

    private static List<WorkingTeam> RankTeams(List<WorkingTeam> rows) =>
        rows.GroupBy(row => row.Event).SelectMany(group => RankEvent(group.ToList())).ToList();

    private static List<WorkingTeam> RankEvent(List<WorkingTeam> rows)
    {
        var ordered = rows
            .OrderByDescending(row => row.Points)
            .ThenByDescending(row => row.TournamentCount)
            .ThenBy(row => row.AveragePlace)
            .ThenBy(row => row.SchoolName, StringComparer.OrdinalIgnoreCase)
            .ToList();
        var ranks = CompetitionRanks(ordered, (left, right) =>
        {
            var points = right.Points.CompareTo(left.Points);
            if (points != 0)
            {
                return points;
            }

            var tournaments = right.TournamentCount.CompareTo(left.TournamentCount);
            if (tournaments != 0)
            {
                return tournaments;
            }

            return left.AveragePlace.CompareTo(right.AveragePlace);
        });
        for (var index = 0; index < ordered.Count; index++)
        {
            ordered[index].Rank = ranks[index];
        }

        return ordered;
    }

    private static List<int> CompetitionRanks<T>(IReadOnlyList<T> ordered, Comparison<T> compare)
    {
        var ranks = new List<int>(ordered.Count);
        for (var index = 0; index < ordered.Count; index++)
        {
            if (index == 0 || compare(ordered[index - 1], ordered[index]) != 0)
            {
                ranks.Add(index + 1);
            }
            else
            {
                ranks.Add(ranks[^1]);
            }
        }

        return ranks;
    }

    private sealed class Appearance
    {
        public string TournamentId { get; init; } = "";
        public string TournamentName { get; init; } = "";
        public DateOnly Start { get; init; }
        public string Week { get; init; } = "";
        public EventKind Event { get; init; }
        public string SchoolId { get; init; } = "";
        public string SchoolName { get; init; } = "";
        public string StudentId { get; init; } = "";
        public string StudentName { get; init; } = "";
        public int? Grade { get; init; }
        public int Points { get; init; }
        public int Place { get; init; }
        public string Detail { get; init; } = "";

        public PointLine ToLine(bool dropped) =>
            new(TournamentId, TournamentName, Start, Points, Place, dropped, Detail);

        public PointLine ToLine(bool dropped, int points, string detail) =>
            new(TournamentId, TournamentName, Start, points, Place, dropped, detail);

        public static Appearance FromEntry(SeasonEntry entry, string unitId, int points, string detail) => new()
        {
            TournamentId = entry.TournamentId,
            TournamentName = entry.TournamentName,
            Start = entry.StartDate,
            Week = SeasonWeek.Key(entry.StartDate),
            Event = entry.Event,
            SchoolId = entry.SchoolId,
            SchoolName = entry.SchoolName,
            StudentId = unitId,
            StudentName = entry.SchoolName,
            Points = points,
            Place = entry.Place,
            Detail = detail
        };

        public static Appearance FromCompetitor(SeasonEntry entry, SeasonCompetitor competitor, int points, string detail) => new()
        {
            TournamentId = entry.TournamentId,
            TournamentName = entry.TournamentName,
            Start = entry.StartDate,
            Week = SeasonWeek.Key(entry.StartDate),
            Event = entry.Event,
            SchoolId = competitor.SchoolId,
            SchoolName = competitor.SchoolName,
            StudentId = competitor.StudentId,
            StudentName = competitor.Name,
            Grade = competitor.Grade,
            Points = points,
            Place = entry.Place,
            Detail = detail
        };
    }

    private sealed record TournamentChoice(Appearance Source, List<Appearance> Top, int Sum);

    private sealed class WorkingIndividual
    {
        public string StudentId { get; set; } = "";
        public string Name { get; set; } = "";
        public string SchoolId { get; set; } = "";
        public string SchoolName { get; set; } = "";
        public EventKind Event { get; set; }
        public int Points { get; set; }
        public int TournamentCount { get; set; }
        public bool QualifiedToState { get; set; }
        public int? Grade { get; set; }
        public double AveragePlace { get; set; }
        public int? AwardRank { get; set; }
        public List<PointLine> Lines { get; set; } = new();

        public IndividualStanding ToStanding() => new(
            StudentId,
            Name,
            SchoolId,
            SchoolName,
            Event,
            Points,
            TournamentCount >= TournamentsRequired,
            TournamentCount,
            QualifiedToState,
            Grade,
            AveragePlace,
            AwardRank,
            Lines);
    }

    private sealed class WorkingTeam
    {
        public string SchoolId { get; set; } = "";
        public string SchoolName { get; set; } = "";
        public EventKind Event { get; set; }
        public int Points { get; set; }
        public int Rank { get; set; }
        public int TournamentCount { get; set; }
        public double AveragePlace { get; set; }
        public List<PointLine> Lines { get; set; } = new();

        public TeamStanding ToStanding() => new(
            SchoolId,
            SchoolName,
            Event,
            Points,
            Rank,
            TournamentCount,
            AveragePlace,
            Lines);
    }
}
