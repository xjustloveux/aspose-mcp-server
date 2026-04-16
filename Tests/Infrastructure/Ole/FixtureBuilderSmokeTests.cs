namespace AsposeMcpServer.Tests.Infrastructure.Ole;

/// <summary>
///     Smoke test for <see cref="FixtureBuilder" />. Proves the 12-fixture matrix builds
///     on the current platform and leaves files on disk; used by the coder stage to
///     satisfy the "fixture generation success" checklist item. Does not exercise the
///     handlers — full AC coverage is added by the test-engineer.
/// </summary>
public sealed class FixtureBuilderSmokeTests
{
    [Fact]
    public void FixtureBuilder_ProducesAllTwelveFixtures()
    {
        using var builder = new FixtureBuilder();

        Assert.Equal(12, builder.Paths.Count);
        foreach (var (kind, path) in builder.Paths)
        {
            Assert.True(File.Exists(path), $"Fixture for {kind} was not written to disk: {path}");
            Assert.True(new FileInfo(path).Length > 0, $"Fixture for {kind} is empty.");
        }
    }
}
