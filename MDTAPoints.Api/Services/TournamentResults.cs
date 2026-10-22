using MDTAPoints.Api.Data;
using MDTAPoints.Core;

namespace MDTAPoints.Api.Services;

public static class TournamentResults
{
    private static readonly EventKind[] EventOrder =
    {
        EventKind.LincolnDouglas,
        EventKind.Classic,
        EventKind.PublicForum,
        EventKind.Policy,
        EventKind.Congress
    };

    public static IReadOnlyList<EventRow> Order(IEnumerable<EventRow> events) =>
        events
            .OrderBy(item => EventIndex(item.EventKind))
            .ThenBy(item => EventTitle(item), StringComparer.OrdinalIgnoreCase)
            .ThenBy(item => item.Id)
            .ToList();

    public static string EventTitle(EventRow item)
    {
        if (!string.IsNullOrWhiteSpace(item.Name))
        {
            return item.Name;
        }

        return Enum.TryParse<EventKind>(item.EventKind, out var kind)
            ? EventKindLabels.Display(kind)
            : item.EventKind;
    }

    public static IReadOnlyList<string> Rules(TournamentDetail tournament, EventRow item)
    {
        if (!TryContext(tournament, item, out var context))
        {
            return new[] { "This event could not be scored." };
        }

        return ScoringGuide.Describe(context);
    }

    public static IReadOnlyList<TournamentResultRow> Rows(TournamentDetail tournament, EventRow item)
    {
        var parsed = TryContext(tournament, item, out var context);
        return item.Entries
            .Select(entry => Row(entry, parsed, context))
            .OrderBy(row => row.Place)
            .ThenBy(row => row.Name, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static TournamentResultRow Row(EntryRow entry, bool parsed, EventScoringContext context)
    {
        var name = People(entry);
        if (!parsed)
        {
            return new TournamentResultRow(entry.Place, name, entry.SchoolName, 0, "This event could not be scored.", entry.PlaceNote);
        }

        var score = CupScorer.Score(new ScoreRequest(
            context.TournamentKind,
            context.Event,
            context.VarsityOrOpen,
            context.InState,
            context.SchoolCount,
            context.EntryCount,
            context.BreaksToOctosOrDeeper,
            context.HasFinalSession,
            context.HasSemis,
            entry.Place,
            entry.ReachedElims,
            entry.ElimWins,
            entry.ReachedFinals,
            entry.ReachedSemis,
            context.IsHouse));
        return new TournamentResultRow(entry.Place, name, entry.SchoolName, score.Points, string.Join(", ", score.Reasons), entry.PlaceNote);
    }

    private static bool TryContext(TournamentDetail tournament, EventRow item, out EventScoringContext context)
    {
        if (!Enum.TryParse<TournamentKind>(tournament.Kind, out var tournamentKind) || !Enum.TryParse<EventKind>(item.EventKind, out var eventKind))
        {
            context = new EventScoringContext(TournamentKind.Invitational, EventKind.LincolnDouglas, false, false, 0, 0, false, false, false, false);
            return false;
        }

        context = new EventScoringContext(
            tournamentKind,
            eventKind,
            item.CountedDivision,
            tournament.InState,
            item.SchoolCount,
            item.EntryCount,
            item.BreaksToOctosOrDeeper,
            item.HasFinalSession,
            item.HasSemis,
            item.IsHouse);
        return true;
    }

    private static string People(EntryRow entry)
    {
        var names = entry.Competitors
            .Select(competitor => competitor.RawName.Trim())
            .Where(name => name.Length > 0)
            .ToArray();
        return names.Length == 0 ? entry.EntryName : string.Join(" & ", names);
    }

    private static int EventIndex(string kind)
    {
        if (!Enum.TryParse<EventKind>(kind, out var parsed))
        {
            return int.MaxValue;
        }

        var index = Array.IndexOf(EventOrder, parsed);
        return index < 0 ? int.MaxValue : index;
    }
}
