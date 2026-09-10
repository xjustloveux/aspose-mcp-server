using AsposeMcpServer.Helpers;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Helpers;

/// <summary>
///     R9-F01: a file a publish could not delete has to be retried, and eventually reported.
///     <para>
///         Announcing it on stderr made the debt visible and nothing more: nothing retried it,
///         nothing aged it out, and a restart forgot it entirely. A superseded document could stay
///         readable on disk with no record that it should not be.
///     </para>
///     <para>
///         The other half of the requirement is that the record never becomes an authority. A path
///         in the queue is a path to consider deleting; whether it may be deleted is decided again
///         at the moment of the attempt, from the filesystem as it then is.
///     </para>
/// </summary>
[Collection("SerialStaticSeams")]
public class CleanupDebtQueueTests : TestBase
{
    /// <summary>A queue backed by a file in this test's own directory.</summary>
    /// <param name="name">Queue file name.</param>
    /// <param name="retention">How long a debt is retried.</param>
    /// <returns>The queue.</returns>
    private CleanupDebtQueue AQueue(string name, TimeSpan? retention = null)
    {
        return new CleanupDebtQueue(CreateTestFilePath(name), [TestDir], Recovery, retention);
    }

    /// <summary>Writes a file for the queue to be asked about.</summary>
    /// <param name="name">File name.</param>
    /// <returns>The file's path.</returns>
    private string AFile(string name)
    {
        var path = CreateTestFilePath(name);
        File.WriteAllText(path, "superseded");
        return path;
    }

    [Fact]
    public void ADeleteThatFailsAndThenSucceeds_ShouldBeRetriedUntilItDoes()
    {
        var queue = AQueue("retry.json");
        var path = AFile("retry_target.txt");

        var attempts = 0;
        queue.DeleteIf = (_, _) =>
        {
            attempts++;
            if (attempts == 1) throw new IOException("the file is locked");
            File.Delete(path);
            return true;
        };

        queue.Record(path, "the file is locked");

        var first = queue.Sweep();
        Assert.Equal([path], first.Retrying);
        Assert.True(File.Exists(path), "the first attempt failed, so the file must still be there");

        // The backoff is real time; move the clock rather than waiting for it.
        queue.Now = () => DateTimeOffset.UtcNow.AddMinutes(10);

        var second = queue.Sweep();
        Assert.Equal([path], second.Deleted);
        Assert.False(File.Exists(path));
        Assert.Empty(queue.Pending());
    }

    [Fact]
    public void ADebtThatKeepsFailingPastItsRetention_ShouldBeAbandonedAndReported()
    {
        var queue = AQueue("retention.json", TimeSpan.FromHours(1));
        var path = AFile("retention_target.txt");

        queue.DeleteIf = (_, _) => throw new IOException("still locked");
        queue.Record(path, "still locked");

        var immediately = queue.Sweep();
        Assert.Equal([path], immediately.Retrying);
        Assert.Empty(immediately.Abandoned);

        queue.Now = () => DateTimeOffset.UtcNow.AddHours(2);

        var later = queue.Sweep();
        Assert.Equal([path], later.Abandoned);
        Assert.Empty(later.Retrying);
        Assert.Empty(queue.Pending());
        Assert.True(File.Exists(path),
            "abandoning a debt means telling someone, not pretending the file is gone");
    }

    [Fact]
    public void ADebtRecordedBeforeARestart_ShouldStillBeThereAfterOne()
    {
        var queueFile = CreateTestFilePath("restart.json");
        var path = AFile("restart_target.txt");

        var before = new CleanupDebtQueue(queueFile, [TestDir], Recovery);
        before.Record(path, "locked at the time");

        // A second instance over the same file is what a restart looks like from here.
        var after = new CleanupDebtQueue(queueFile, [TestDir], Recovery);

        var pending = after.Pending();
        Assert.Single(pending);
        Assert.Equal(path, pending[0].Path);

        var swept = after.Sweep();
        Assert.Equal([path], swept.Deleted);
        Assert.False(File.Exists(path));
    }

    [SkippableFact]
    public void APathThatHasBecomeALink_ShouldBeRefusedRatherThanFollowed()
    {
        SkipIfNotWindows("Creating a file link requires privileges this account may not have.");

        var queueFile = CreateTestFilePath("link.json");
        var real = AFile("link_real_target.txt");
        var elsewhere = AFile("link_elsewhere.txt");

        var queue = new CleanupDebtQueue(queueFile, [TestDir], Recovery);
        queue.Record(real, "locked at the time");

        // Probed before the assertion, so an account without the privilege skips rather than
        // failing on a SkipException raised from inside a catch.
        var probe = CreateTestFilePath("link_probe.txt");
        var canLink = true;
        try
        {
            File.CreateSymbolicLink(probe, elsewhere);
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
        {
            canLink = false;
        }

        Skip.If(!canLink, "This account cannot create symbolic links.");

        File.Delete(real);
        File.CreateSymbolicLink(real, elsewhere);

        var swept = queue.Sweep();

        Assert.Equal([real], swept.Refused);
        Assert.Empty(swept.Deleted);
        Assert.True(File.Exists(elsewhere), "the link's target must not have been deleted");
        Assert.Empty(queue.Pending());
    }

    [Fact]
    public void APathOutsideTheAllowlist_ShouldNeverBeDeleted()
    {
        // The queue is written from inside this process, but a queue file is still a file: if a
        // path in it were obeyed without checking, the record would be a deletion primitive.
        var outside = Path.Combine(Path.GetTempPath(),
            "cleanup_outside_" + Guid.NewGuid().ToString("N") + ".txt");
        File.WriteAllText(outside, "not ours to delete");

        try
        {
            var queue = new CleanupDebtQueue(CreateTestFilePath("outside.json"), [TestDir], Recovery);
            queue.Record(outside, "locked at the time");

            var swept = queue.Sweep();

            Assert.Equal([Path.GetFullPath(outside)], swept.Refused);
            Assert.True(File.Exists(outside));
        }
        finally
        {
            File.Delete(outside);
        }
    }

    [Fact]
    public void AFileSomeoneElseRemoved_ShouldLeaveTheQueue()
    {
        var queue = AQueue("gone.json");
        var path = AFile("gone_target.txt");

        queue.Record(path, "locked at the time");
        File.Delete(path);

        var swept = queue.Sweep();

        Assert.Equal([path], swept.Deleted);
        Assert.Empty(queue.Pending());
    }

    [Fact]
    public void ABatchThatCannotCleanUp_ShouldHandTheDebtToWhateverIsListening()
    {
        // The wiring, not the queue: a batch that announces a debt and tells nobody who can retry
        // it is what this whole mechanism was before (R9-F01).
        List<(string Path, string Error)> recorded = [];
        var previous = BoundedFileBatch.RecordDebt;
        BoundedFileBatch.RecordDebt = (path, error) => recorded.Add((path, error));

        try
        {
            var destination = CreateTestFilePath("handoff.txt");
            using (var batch = new BoundedFileBatch(1024, "a fixture", Recovery, [TestDir]))
            {
                batch.ReportDebt = _ => { };
                batch.BeforeDelete = _ => throw new IOException("the staging file is locked");
                batch.Stage(destination, stream => stream.Write("content"u8));
            }

            Assert.NotEmpty(recorded);
            Assert.All(recorded, entry => Assert.False(string.IsNullOrWhiteSpace(entry.Path)));
        }
        finally
        {
            BoundedFileBatch.RecordDebt = previous;
        }
    }

    [Fact]
    public void RecordingTheSamePathTwice_ShouldNotRestartItsRetention()
    {
        // Otherwise a path that fails on every publish is never old enough to be abandoned.
        var queue = AQueue("age.json", TimeSpan.FromHours(1));
        var path = AFile("age_target.txt");

        queue.DeleteIf = (_, _) => throw new IOException("still locked");

        var start = DateTimeOffset.UtcNow;
        queue.Now = () => start;
        queue.Record(path, "first failure");

        queue.Now = () => start.AddMinutes(50);
        queue.Record(path, "failed again");

        queue.Now = () => start.AddMinutes(70);
        var swept = queue.Sweep();

        Assert.Equal([path], swept.Abandoned);
    }
}
