namespace MDTAPoints.Tabroom;

public sealed class TabroomClient
{
    public const string DefaultBaseUrl = "https://api.tabroom.com/v1/rest/";

    private readonly HttpClient _http;

    public TabroomClient(HttpClient http)
    {
        _http = http;
    }

    public async Task<string> GetTournamentJsonAsync(int tournamentId, CancellationToken cancellationToken) =>
        await _http.GetStringAsync(TournamentPath(tournamentId), cancellationToken);

    public async Task<string> GetEventsJsonAsync(int tournamentId, CancellationToken cancellationToken) =>
        await _http.GetStringAsync($"{TournamentPath(tournamentId)}/results", cancellationToken);

    public async Task<string> GetResultJsonAsync(int tournamentId, int resultSetId, CancellationToken cancellationToken) =>
        await _http.GetStringAsync($"{TournamentPath(tournamentId)}/results/{resultSetId}", cancellationToken);

    private static string TournamentPath(int tournamentId) => $"tourns/{tournamentId}";
}
