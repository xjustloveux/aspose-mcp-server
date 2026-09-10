using AsposeMcpServer.Helpers;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Helpers;

/// <summary>
///     R9-F01: the debt this whole mechanism exists for is the one that has to reach the reaper.
///     <para>
///         When a publish replaces an existing file, the previous version is moved aside to a
///         <c>.replaced-*</c> sibling and deleted once every output is in place. A delete that
///         fails there leaves the caller's *previous document* on disk. That failure produced a
///         human-readable line in <c>CleanupFailures</c> and nothing else — the path was never
///         added to the batch's debts, so it was announced once and never retried, while the
///         staging-cleanup failures beside it were queued properly.
///     </para>
///     <para>
///         The earlier handoff fixture never called <see cref="BoundedFileBatch.Publish" />: it
///         failed a staging delete, which took the path that already worked. This one drives the
///         whole sequence — existing destination, successful publish, a delete that fails only for
///         the backup — and follows the path all the way into the queue and out again.
///     </para>
/// </summary>
[Collection("SerialStaticSeams")]
public class BackupDebtHandoffTests : TestBase
{
    /// <summary>A queue writing its state beside the fixture files.</summary>
    /// <param name="name">Queue file name.</param>
    /// <returns>The queue.</returns>
    private CleanupDebtQueue AQueue(string name)
    {
        return new CleanupDebtQueue(CreateTestFilePath(name), [TestDir], Recovery, CleanupDebtQueue.DefaultRetention);
    }

    /// <summary>
    ///     Publishes over an existing file with the backup delete failing, and reports what the
    ///     batch handed to <see cref="BoundedFileBatch.RecordDebt" />.
    /// </summary>
    /// <param name="destination">The file to replace.</param>
    /// <param name="recorded">Collects what was recorded.</param>
    /// <returns>The backup path the publish created.</returns>
    private string PublishOverAnExistingFileWithAFailingBackupDelete(
        string destination, List<(string Path, string Error)> recorded)
    {
        string? backup = null;
        var previous = BoundedFileBatch.RecordDebt;
        BoundedFileBatch.RecordDebt = (path, error) => recorded.Add((path, error));

        try
        {
            using var batch = new BoundedFileBatch(4096, "a fixture", Recovery, [Path.GetDirectoryName(destination)!]);
            batch.ReportDebt = _ => { };

            // Only the backup delete fails. A blanket failure would also fail the staging cleanup,
            // which already worked, and the test would pass on the wrong path.
            batch.BeforeDelete = path =>
            {
                if (!path.Contains(".replaced-", StringComparison.Ordinal)) return;

                backup = path;
                throw new IOException("the previous version is locked");
            };

            batch.Stage(destination, stream => stream.Write("new content"u8));
            batch.Publish();
        }
        finally
        {
            BoundedFileBatch.RecordDebt = previous;
        }

        Assert.NotNull(backup);
        return backup!;
    }

    [Fact]
    public void ABackupThatCouldNotBeDeleted_ShouldReachTheReaperByItsFullPath()
    {
        var destination = CreateTestFilePath("backup_handoff.txt");
        File.WriteAllText(destination, "previous content");

        var recorded = new List<(string Path, string Error)>();
        var backup = PublishOverAnExistingFileWithAFailingBackupDelete(destination, recorded);

        // The request itself succeeded: the caller has the new document.
        Assert.Equal("new content", File.ReadAllText(destination));
        Assert.True(File.Exists(backup), "the fixture did not actually leave a backup behind");

        // And the *path*, not a sentence naming the file, is what the reaper needs.
        Assert.Contains(recorded, entry => entry.Path == backup);
    }

    [Fact]
    public void OnlyTheBackup_ShouldBeQueuedWhenTheRestOfTheCleanupSucceeded()
    {
        var destination = CreateTestFilePath("backup_only.txt");
        File.WriteAllText(destination, "previous content");

        var recorded = new List<(string Path, string Error)>();
        var backup = PublishOverAnExistingFileWithAFailingBackupDelete(destination, recorded);

        Assert.Equal([backup], recorded.Select(entry => entry.Path).ToList());
    }

    [Fact]
    public void AQueuedBackup_ShouldBeDeletedByTheNextSweep()
    {
        var destination = CreateTestFilePath("backup_swept.txt");
        File.WriteAllText(destination, "previous content");

        var recorded = new List<(string Path, string Error)>();
        var backup = PublishOverAnExistingFileWithAFailingBackupDelete(destination, recorded);

        var queue = AQueue("backup_swept_queue.json");
        foreach (var (path, error) in recorded) queue.Record(path, error);

        Assert.Contains(backup, queue.Pending().Select(debt => debt.Path));

        var swept = queue.Sweep();

        Assert.Contains(backup, swept.Deleted);
        Assert.False(File.Exists(backup), "the sweep reported the backup removed but it is still there");
        Assert.Equal("new content", File.ReadAllText(destination));
    }

    [Fact]
    public void AQueuedBackup_ShouldSurviveARestart()
    {
        // The queue is only worth having if a debt outlives the process that recorded it: a
        // backup left behind by a crash is exactly the case nobody is watching.
        var destination = CreateTestFilePath("backup_restart.txt");
        File.WriteAllText(destination, "previous content");

        var recorded = new List<(string Path, string Error)>();
        var backup = PublishOverAnExistingFileWithAFailingBackupDelete(destination, recorded);

        var queueFile = CreateTestFilePath("backup_restart_queue.json");
        var first = new CleanupDebtQueue(queueFile, [TestDir], Recovery, CleanupDebtQueue.DefaultRetention);
        foreach (var (path, error) in recorded) first.Record(path, error);

        var afterRestart = new CleanupDebtQueue(queueFile, [TestDir], Recovery, CleanupDebtQueue.DefaultRetention);

        Assert.Contains(backup, afterRestart.Pending().Select(debt => debt.Path));
        Assert.Contains(backup, afterRestart.Sweep().Deleted);
        Assert.False(File.Exists(backup));
    }

    [Fact]
    public void APublishWhoseCleanupSucceeded_ShouldQueueNothing()
    {
        // The control. Without it a mechanism that queued every publish would pass every test
        // above and fill the queue with paths that were deleted successfully.
        var destination = CreateTestFilePath("backup_clean.txt");
        File.WriteAllText(destination, "previous content");

        var recorded = new List<(string Path, string Error)>();
        var previous = BoundedFileBatch.RecordDebt;
        BoundedFileBatch.RecordDebt = (path, error) => recorded.Add((path, error));

        try
        {
            using var batch = new BoundedFileBatch(4096, "a fixture", Recovery, [TestDir]);
            batch.Stage(destination, stream => stream.Write("new content"u8));
            batch.Publish();
        }
        finally
        {
            BoundedFileBatch.RecordDebt = previous;
        }

        Assert.Empty(recorded);
        Assert.Equal("new content", File.ReadAllText(destination));
        Assert.Empty(Directory.GetFiles(TestDir, "backup_clean.txt.replaced-*"));
    }
}
