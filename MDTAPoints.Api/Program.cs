using MDTAPoints.Api.Data;
using MDTAPoints.Api.Services;
using MDTAPoints.Core;
using MDTAPoints.Tabroom;
using Microsoft.AspNetCore.Authentication.Cookies;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddRazorPages(options =>
{
    options.Conventions.AuthorizeFolder("/Admin");
});
builder.Services.AddAuthentication(CookieAuthenticationDefaults.AuthenticationScheme)
    .AddCookie(options =>
    {
        options.LoginPath = "/Login";
        options.LogoutPath = "/Logout";
        options.Cookie.HttpOnly = true;
        options.Cookie.SameSite = SameSiteMode.Lax;
        options.Cookie.SecurePolicy = CookieSecurePolicy.SameAsRequest;
        options.ExpireTimeSpan = TimeSpan.FromHours(8);
        options.SlidingExpiration = true;
    });
builder.Services.AddAuthorization();
builder.Services.AddSingleton<SchemaMigrator>();
builder.Services.AddScoped<CupDatabase>();
builder.Services.AddHttpClient<TabroomClient>(client =>
{
    client.BaseAddress = new Uri(TabroomClient.DefaultBaseUrl);
    client.Timeout = TimeSpan.FromSeconds(60);
    client.DefaultRequestHeaders.UserAgent.ParseAdd("MDTAPoints/1.0");
});
builder.Services.AddScoped<TabroomImportService>();

var app = builder.Build();

using (var scope = app.Services.CreateScope())
{
    var migrator = scope.ServiceProvider.GetRequiredService<SchemaMigrator>();
    await migrator.ApplyAsync(CancellationToken.None);
    var database = scope.ServiceProvider.GetRequiredService<CupDatabase>();
    await database.SeedAdminAsync(app.Configuration, CancellationToken.None);
}

if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Error");
    app.UseHsts();
}

app.UseHttpsRedirection();
app.UseStaticFiles();
app.UseRouting();
app.UseAuthentication();
app.UseAuthorization();

app.MapGet("/api/standings/individuals", async (string? eventKind, CupDatabase database, CancellationToken cancellationToken) =>
{
    var standings = StandingsCalculator.Calculate(await database.LoadSeasonAsync(cancellationToken));
    var rows = standings.Individuals.AsEnumerable();
    if (Enum.TryParse<EventKind>(eventKind, true, out var kind))
    {
        rows = rows.Where(row => row.Event == kind);
    }

    return Results.Ok(rows);
});
app.MapGet("/api/standings/teams", async (string? eventKind, CupDatabase database, CancellationToken cancellationToken) =>
{
    var standings = StandingsCalculator.Calculate(await database.LoadSeasonAsync(cancellationToken));
    var rows = standings.Teams.AsEnumerable();
    if (Enum.TryParse<EventKind>(eventKind, true, out var kind))
    {
        rows = rows.Where(row => row.Event == kind);
    }

    return Results.Ok(rows);
});

var adminApi = app.MapGroup("/api").RequireAuthorization();
adminApi.MapGet("/tournaments", async (CupDatabase database, CancellationToken cancellationToken) =>
    Results.Ok(await database.ListTournamentsAsync(cancellationToken)));
adminApi.MapGet("/tournaments/{id:int}", async (int id, CupDatabase database, CancellationToken cancellationToken) =>
{
    var tournament = await database.GetTournamentAsync(id, cancellationToken);
    return tournament is null ? Results.NotFound() : Results.Ok(tournament);
});
adminApi.MapPost("/imports/tabroom", async (TabroomImportRequest request, TabroomImportService imports, CancellationToken cancellationToken) =>
{
    var outcome = await imports.ImportAsync(request.TournamentId, cancellationToken);
    return Results.Ok(outcome);
});
adminApi.MapPost("/imports/{id:int}/publish", async (int id, CupDatabase database, CancellationToken cancellationToken) =>
{
    try
    {
        await database.PublishAsync(id, cancellationToken);
        return Results.Ok();
    }
    catch (InvalidOperationException ex)
    {
        return Results.BadRequest(new { error = ex.Message });
    }
});
adminApi.MapPost("/tournaments", async (CreateTournamentRequest request, CupDatabase database, CancellationToken cancellationToken) =>
{
    var id = await database.CreateManualAsync(request.Name, request.StartDate, request.State, request.Kind, cancellationToken);
    return Results.Ok(new { id });
});
adminApi.MapPost("/tournaments/{id:int}/entries", async (int id, CreateEntryRequest request, CupDatabase database, CancellationToken cancellationToken) =>
{
    await database.AddEntryAsync(id, request.EventKind, request.EventName, request.Place, request.School, request.Competitors, request.Prelim, request.ElimWins, request.ReachedFinals, request.ReachedSemis, request.QualifiedToState, cancellationToken);
    return Results.Ok();
});
adminApi.MapPatch("/tournaments/{id:int}", async (int id, PatchTournamentRequest request, CupDatabase database, CancellationToken cancellationToken) =>
{
    var update = new ReviewUpdate { Kind = request.Kind ?? TournamentKind.Invitational.ToString() };
    await database.SaveReviewAsync(id, update, cancellationToken);
    return Results.Ok();
});
adminApi.MapPatch("/students/{id:int}", async (int id, PatchStudentRequest request, CupDatabase database, CancellationToken cancellationToken) =>
{
    await database.UpdateStudentAsync(id, request.Grade, request.QualifiedEvents ?? new List<string>(), cancellationToken);
    return Results.Ok();
});

app.MapRazorPages();
app.Run();

public sealed record TabroomImportRequest(int TournamentId);
public sealed record CreateTournamentRequest(string Name, DateOnly StartDate, string? State, string Kind);
public sealed record CreateEntryRequest(string EventKind, string EventName, int Place, string School, string Competitors, bool Prelim, int ElimWins, bool ReachedFinals, bool ReachedSemis, bool QualifiedToState);
public sealed record PatchTournamentRequest(string? Kind);
public sealed record PatchStudentRequest(int? Grade, List<string>? QualifiedEvents);
