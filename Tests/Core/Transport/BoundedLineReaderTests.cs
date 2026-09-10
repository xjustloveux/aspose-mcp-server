using AsposeMcpServer.Core.Transport;

namespace AsposeMcpServer.Tests.Core.Transport;

/// <summary>
///     The WebSocket bridge reads the child process line by line. <c>ReadLineAsync</c> has no
///     bound, so a child that writes without ever writing a newline made the server accumulate
///     everything it emitted (R3-R02). These fixtures drive the bounded reader that replaced it.
/// </summary>
public class BoundedLineReaderTests
{
    /// <summary>
    ///     Reads every line the reader yields for the given text.
    /// </summary>
    /// <param name="text">Text to read.</param>
    /// <param name="maxChars">Line bound.</param>
    /// <returns>The lines, and whether the bound was passed on the last one read.</returns>
    private static async Task<(List<string> Lines, bool TooLong)> ReadAll(string text, int maxChars)
    {
        var reader = new BoundedLineReader(new StringReader(text), maxChars);
        List<string> lines = [];
        var tooLong = false;

        while (await reader.ReadLineAsync(CancellationToken.None) is { } line)
        {
            lines.Add(line);
            if (!reader.LineTooLong) continue;
            tooLong = true;
            break;
        }

        return (lines, tooLong);
    }

    [Fact]
    public async Task ReadLineAsync_ShouldSplitOnNewlines()
    {
        var (lines, tooLong) = await ReadAll("first\nsecond\nthird\n", 100);

        Assert.Equal(["first", "second", "third"], lines);
        Assert.False(tooLong);
    }

    [Fact]
    public async Task ReadLineAsync_ShouldDropTheCarriageReturn()
    {
        var (lines, _) = await ReadAll("first\r\nsecond\r\n", 100);

        Assert.Equal(["first", "second"], lines);
    }

    [Fact]
    public async Task ReadLineAsync_WithNoTrailingNewline_ShouldStillReturnTheLastLine()
    {
        var (lines, _) = await ReadAll("first\nlast", 100);

        Assert.Equal(["first", "last"], lines);
    }

    [Fact]
    public async Task ReadLineAsync_WithAnEmptyStream_ShouldReturnNothing()
    {
        var (lines, _) = await ReadAll("", 100);

        Assert.Empty(lines);
    }

    [Fact]
    public async Task ReadLineAsync_AtExactlyTheBound_ShouldBeAccepted()
    {
        var (lines, tooLong) = await ReadAll(new string('a', 64) + "\n", 64);

        Assert.Equal([new string('a', 64)], lines);
        Assert.False(tooLong);
    }

    /// <summary>
    ///     One character past the bound is the case the unbounded reader could not see.
    /// </summary>
    [Fact]
    public async Task ReadLineAsync_OnePastTheBound_ShouldReportIt()
    {
        var (lines, tooLong) = await ReadAll(new string('a', 65) + "\n", 64);

        Assert.True(tooLong);
        Assert.Equal(64, lines[^1].Length);
    }

    /// <summary>
    ///     A line far larger than the internal buffer must be stopped at the bound rather than
    ///     accumulated, which is the failure this reader exists to prevent.
    /// </summary>
    [Fact]
    public async Task ReadLineAsync_WithNoNewlineAtAll_ShouldStopAtTheBound()
    {
        var (lines, tooLong) = await ReadAll(new string('x', 500_000), 1_024);

        Assert.True(tooLong);
        Assert.Equal(1_024, lines[^1].Length);
    }

    [Fact]
    public async Task ReadLineAsync_ShouldReadLinesLongerThanItsBuffer()
    {
        var long1 = new string('a', 20_000);
        var long2 = new string('b', 9_000);

        var (lines, tooLong) = await ReadAll(long1 + "\n" + long2 + "\n", 100_000);

        Assert.Equal([long1, long2], lines);
        Assert.False(tooLong);
    }
}
