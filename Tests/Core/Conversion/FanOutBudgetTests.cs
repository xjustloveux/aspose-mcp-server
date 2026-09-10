using System.Text;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Core.Conversion;

/// <summary>
///     R8-C03: a conversion that writes one file per page is one request, with one budget.
///     <para>
///         Each page used to be published on its own, so every page started from the full byte
///         limit — a hundred pages could produce a hundred times what the limit says a request may
///         produce — and each file was moved to its destination as it was made, so a failure part
///         way through left the earlier pages behind with no way to tell a complete result from a
///         truncated one.
///     </para>
///     <para>
///         These fixtures exercise <see cref="BoundedFileBatch" /> directly at the sizes involved.
///         Driving the converter itself would mean rendering pages large enough to exhaust a 2 GiB
///         budget, which is the reason the original bound was never tested.
///     </para>
/// </summary>
public class FanOutBudgetTests : TestBase
{
    /// <summary>Writes a given number of bytes to a stream.</summary>
    /// <param name="count">How many bytes to write.</param>
    /// <returns>The write action.</returns>
    private static Action<Stream> Writing(int count)
    {
        return stream => stream.Write(new byte[count]);
    }

    [Fact]
    public void EveryPageOfAFanOut_ShouldDrawOnTheSameBudget()
    {
        // Three pages of 40 bytes against a 100-byte request budget: the third does not fit, and
        // it is only the shared accounting that knows so.
        using var batch = new BoundedFileBatch(100, "rendered pages", Recovery, []);

        batch.Stage(Path.Combine(TestDir, "page_1.bin"), Writing(40));
        batch.Stage(Path.Combine(TestDir, "page_2.bin"), Writing(40));

        var exhausted = Assert.Throws<ArgumentException>(() =>
            batch.Stage(Path.Combine(TestDir, "page_3.bin"), Writing(40)));

        Assert.Contains("rendered pages", exhausted.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void APerPageBudget_WouldHaveLetTheSameThreePagesThrough()
    {
        // The contrast: with a budget per page — which is what publishing each page separately
        // gave every one of them — all three fit, three times over.
        foreach (var page in new[] { "solo_1.bin", "solo_2.bin", "solo_3.bin" })
            BoundedFilePublisher.Publish(Path.Combine(TestDir, page), 100, Writing(40), "page", Recovery, []);

        Assert.Equal(3, Directory.GetFiles(TestDir, "solo_*.bin").Length);
        Assert.Equal(120, Directory.GetFiles(TestDir, "solo_*.bin").Sum(f => new FileInfo(f).Length));
    }

    [Fact]
    public void AFailureOnTheLastPage_ShouldLeaveNoEarlierPageBehind()
    {
        using var batch = new BoundedFileBatch(1_000, "rendered pages", Recovery, []);

        batch.Stage(Path.Combine(TestDir, "keep_1.bin"), Writing(10));
        batch.Stage(Path.Combine(TestDir, "keep_2.bin"), Writing(10));

        Assert.ThrowsAny<Exception>(() =>
            batch.Stage(Path.Combine(TestDir, "keep_3.bin"),
                _ => throw new InvalidOperationException("the renderer failed")));

        // Nothing was published, so none of the destinations exists — not even the two pages that
        // were produced successfully.
        Assert.Empty(Directory.GetFiles(TestDir, "keep_*.bin"));
    }

    [Fact]
    public void ACompleteFanOut_ShouldPublishEveryPageAtOnce()
    {
        using var batch = new BoundedFileBatch(1_000, "rendered pages", Recovery, []);

        batch.Stage(Path.Combine(TestDir, "done_1.bin"), Writing(10));
        batch.Stage(Path.Combine(TestDir, "done_2.bin"), Writing(10));

        Assert.Empty(Directory.GetFiles(TestDir, "done_*.bin"));

        batch.Publish();

        Assert.Equal(2, Directory.GetFiles(TestDir, "done_*.bin").Length);
        Assert.Empty(Directory.GetFiles(TestDir, "*.partial-*"));
    }

    [Fact]
    public void TheConverterFanOuts_ShouldStageIntoOneBatch()
    {
        // The three loops that write one file per page or per sheet each hold a single batch, so
        // the properties above apply to them rather than to a helper nobody calls.
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory != null && !File.Exists(Path.Combine(directory.FullName, "AsposeMcpServer.csproj")))
            directory = directory.Parent;

        Assert.NotNull(directory);
        var source = File.ReadAllText(
            Path.Combine(directory.FullName, "Core", "Conversion", "DocumentConverter.cs"),
            Encoding.UTF8);

        Assert.Equal(3, source.Split("new BoundedFileBatch(").Length - 1);
        Assert.DoesNotContain("BoundedFilePublisher.Publish(resolvedPagePath", source,
            StringComparison.Ordinal);
        Assert.DoesNotContain("BoundedFilePublisher.Publish(resolvedSheetPath", source,
            StringComparison.Ordinal);
    }
}
