using MDTAPoints.Core;
using MDTAPoints.Tabroom;

if (args is not ["score", var idText] || !int.TryParse(idText, out var tournamentId))
{
    Console.WriteLine("Usage: dotnet run --project MDTAPoints -- score <tabroomTournamentId>");
    Console.WriteLine("Prints Cup points for one tournament. The season, name links, and standings live in the web app.");
    return;
}

using var http = new HttpClient
{
    BaseAddress = new Uri(TabroomClient.DefaultBaseUrl),
    Timeout = TimeSpan.FromSeconds(60)
};
http.DefaultRequestHeaders.UserAgent.ParseAdd("MDTAPoints/1.0");
var client = new TabroomClient(http);
var info = TabroomParser.ParseTournament(await client.GetTournamentJsonAsync(tournamentId, CancellationToken.None));
var summaries = TabroomParser.ParseEvents(await client.GetEventsJsonAsync(tournamentId, CancellationToken.None));

Console.WriteLine($"{info.Name}  {info.StartDate:yyyy-MM-dd}  {info.SuggestedKind}");
foreach (var summary in summaries.Where(item => item.ShouldImport && item.ResultSetIds.Count > 0))
{
    var resultJson = new List<string>();
    foreach (var setId in summary.ResultSetIds)
    {
        resultJson.Add(await client.GetResultJsonAsync(tournamentId, setId, CancellationToken.None));
    }

    var parsed = TabroomParser.ParseResultSets(resultJson);
    var kind = Enum.Parse<TournamentKind>(info.SuggestedKind);
    var eventKind = Enum.Parse<EventKind>(parsed.EventKind);
    Console.WriteLine();
    Console.WriteLine($"{parsed.Name}  {parsed.SchoolCount} schools, {parsed.EntryCount} entries");
    foreach (var entry in parsed.Entries.Where(item => TabroomParser.IncludeSchool(item.SchoolState)).OrderBy(item => item.Place))
    {
        var score = CupScorer.Score(new ScoreRequest(
            kind,
            eventKind,
            parsed.CountedDivision,
            info.InState,
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
        Console.WriteLine($"{entry.Place,4}  {score.Points,2}  {entry.EntryName}  {entry.SchoolName}");
    }
}
