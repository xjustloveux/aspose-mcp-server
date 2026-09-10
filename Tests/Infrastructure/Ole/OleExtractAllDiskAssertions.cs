using AsposeMcpServer.Results.Shared.Ole;

namespace AsposeMcpServer.Tests.Infrastructure.Ole;

/// <summary>
///     Checks what the cumulative byte cap on <c>extract_all</c> actually left on disk.
///     <para>
///         The three size-cap suites asserted on the result object alone — <c>Truncated</c>, the
///         skip reason and the extracted count — so a handler that reported a stop while still
///         writing every payload, or that left a half-written file behind at the point it stopped,
///         passed all of them. The cap exists to bound bytes on disk, and that is the thing worth
///         measuring (R7-T01).
///     </para>
/// </summary>
internal static class OleExtractAllDiskAssertions
{
    /// <summary>
    ///     Asserts that the files in the output directory match what the handler reported and stay
    ///     within the cap.
    /// </summary>
    /// <param name="outputDirectory">Directory the handler extracted into.</param>
    /// <param name="sourceDocument">The fixture document, which lives in the same directory.</param>
    /// <param name="data">The handler's own report.</param>
    /// <param name="cap">The configured cumulative byte cap.</param>
    internal static void AssertDiskAgreesWithReport(
        string outputDirectory, string sourceDocument, OleExtractAllResult data, long cap)
    {
        var written = Directory.GetFiles(outputDirectory)
            .Where(f => !string.Equals(f, sourceDocument, StringComparison.OrdinalIgnoreCase))
            .OrderBy(f => f, StringComparer.Ordinal)
            .ToList();

        Assert.True(written.Count == data.Extracted,
            $"The report claims {data.Extracted} extracted object(s) but {written.Count} file(s) "
            + $"were written: {string.Join(", ", written.Select(Path.GetFileName))}");

        var total = written.Sum(f => new FileInfo(f).Length);
        Assert.True(total <= cap,
            $"{total} byte(s) were written with a cap of {cap}; the cap bounds disk, not the "
            + "summary the handler prints.");

        var leftovers = written
            .Where(f => new FileInfo(f).Length == 0
                        || Path.GetFileName(f).Contains(".partial", StringComparison.Ordinal))
            .Select(Path.GetFileName)
            .ToList();

        Assert.True(leftovers.Count == 0,
            "Stopping at the cap left a partial or empty file behind: "
            + string.Join(", ", leftovers));
    }
}
