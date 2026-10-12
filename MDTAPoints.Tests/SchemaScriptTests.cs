using MDTAPoints.Api.Data;

namespace MDTAPoints.Tests;

public class SchemaScriptTests
{
    [Fact]
    public void Startup_script_is_embedded_and_creates_the_season_tables()
    {
        var assembly = typeof(SchemaMigrator).Assembly;
        var name = Assert.Single(assembly.GetManifestResourceNames(), resource => resource.EndsWith(".sql", StringComparison.OrdinalIgnoreCase));
        Assert.Equal("001_create", name.Split('.')[^2]);

        using var stream = assembly.GetManifestResourceStream(name);
        Assert.NotNull(stream);
        using var reader = new StreamReader(stream);
        var sql = reader.ReadToEnd();
        Assert.Contains("CREATE TABLE dbo.Students", sql, StringComparison.Ordinal);
        Assert.Contains("CREATE TABLE dbo.Tournaments", sql, StringComparison.Ordinal);
        Assert.Contains("CREATE TABLE dbo.Users", sql, StringComparison.Ordinal);
        Assert.Contains("CREATE TABLE dbo.SchemaVersion", sql, StringComparison.Ordinal);
    }
}
