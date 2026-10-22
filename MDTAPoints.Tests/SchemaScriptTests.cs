using MDTAPoints.Api.Data;

namespace MDTAPoints.Tests;

public class SchemaScriptTests
{
    [Fact]
    public void Startup_script_is_embedded_and_creates_the_season_tables()
    {
        var assembly = typeof(SchemaMigrator).Assembly;
        var names = assembly.GetManifestResourceNames()
            .Where(resource => resource.EndsWith(".sql", StringComparison.OrdinalIgnoreCase))
            .OrderBy(resource => resource, StringComparer.Ordinal)
            .ToArray();
        Assert.Equal(2, names.Length);
        Assert.Equal("001_create", names[0].Split('.')[^2]);
        Assert.Equal("002_entry_place", names[1].Split('.')[^2]);

        using var stream = assembly.GetManifestResourceStream(names[0]);
        Assert.NotNull(stream);
        using var reader = new StreamReader(stream);
        var sql = reader.ReadToEnd();
        Assert.Contains("CREATE TABLE dbo.Students", sql, StringComparison.Ordinal);
        Assert.Contains("CREATE TABLE dbo.Tournaments", sql, StringComparison.Ordinal);
        Assert.Contains("CREATE TABLE dbo.Users", sql, StringComparison.Ordinal);
        Assert.Contains("CREATE TABLE dbo.SchemaVersion", sql, StringComparison.Ordinal);

        using var placeStream = assembly.GetManifestResourceStream(names[1]);
        Assert.NotNull(placeStream);
        using var placeReader = new StreamReader(placeStream);
        var placeSql = placeReader.ReadToEnd();
        Assert.Contains("TabroomPlace", placeSql, StringComparison.Ordinal);
        Assert.Contains("PlaceNote", placeSql, StringComparison.Ordinal);
    }
}
