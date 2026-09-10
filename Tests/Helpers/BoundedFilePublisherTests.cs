using AsposeMcpServer.Helpers;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Helpers;

/// <summary>
///     Extraction wrote each part straight onto the destination and measured it afterwards, so the
///     file that passed the limit was already on disk when the request was refused (R3-R07).
/// </summary>
public class BoundedFilePublisherTests : TestBase
{
    [Fact]
    public void Publish_WithinTheBudget_ShouldWriteTheDestination()
    {
        var destination = CreateTestFilePath("published.txt");

        var written = BoundedFilePublisher.Publish(destination, 1_000,
            stream => stream.Write("hello"u8), "output", Recovery, []);

        Assert.Equal(5, written.WrittenBytes);
        Assert.Equal("hello", File.ReadAllText(destination));
    }

    [Fact]
    public void Publish_PastTheBudget_ShouldLeaveNoDestination()
    {
        var destination = CreateTestFilePath("refused.txt");

        Assert.Throws<ArgumentException>(() => BoundedFilePublisher.Publish(destination, 4,
            stream => stream.Write(new byte[5]), "output", Recovery, []));

        Assert.False(File.Exists(destination));
    }

    /// <summary>
    ///     A refusal must not damage what the caller already had at that path.
    /// </summary>
    [Fact]
    public void Publish_PastTheBudget_ShouldLeaveAnExistingFileUntouched()
    {
        var destination = CreateTestFilePath("existing.txt");
        File.WriteAllText(destination, "original");

        Assert.Throws<ArgumentException>(() => BoundedFilePublisher.Publish(destination, 4,
            stream => stream.Write(new byte[5]), "output", Recovery, []));

        Assert.Equal("original", File.ReadAllText(destination));
    }

    [Fact]
    public void Publish_PastTheBudget_ShouldLeaveNoStagingFileBehind()
    {
        var destination = CreateTestFilePath("staging_cleanup.txt");

        Assert.Throws<ArgumentException>(() => BoundedFilePublisher.Publish(destination, 4,
            stream => stream.Write(new byte[5]), "output", Recovery, []));

        Assert.Empty(Directory.GetFiles(TestDir, "staging_cleanup.txt.partial-*"));
    }

    [Fact]
    public void Publish_WithNoBudgetLeft_ShouldRefuseBeforeWritingAnything()
    {
        var destination = CreateTestFilePath("no_budget.txt");
        var called = false;

        Assert.Throws<ArgumentException>(() => BoundedFilePublisher.Publish(destination, 0,
            _ => called = true, "output", Recovery, []));

        Assert.False(called);
        Assert.False(File.Exists(destination));
    }

    [Fact]
    public void Publish_ShouldOverwriteAnExistingDestination()
    {
        var destination = CreateTestFilePath("overwrite.txt");
        File.WriteAllText(destination, "old content that is longer");

        BoundedFilePublisher.Publish(destination, 1_000,
            stream => stream.Write("new"u8), "output", Recovery, []);

        Assert.Equal("new", File.ReadAllText(destination));
    }
}
