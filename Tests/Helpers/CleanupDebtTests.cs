using AsposeMcpServer.Helpers;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Helpers;

/// <summary>
///     R8-F01 and R8-F02: what a publish cannot tidy up has to be sayable, and saying it must not
///     replace the failure that caused it.
///     <para>
///         The staging cleanup went straight to <c>File.Delete</c> with no re-resolution and no
///         seam, so it could not be exercised at all, and an error from it replaced the original
///         failure. After a commit, a backup that could not be removed was recorded in
///         <c>CleanupFailures</c> — which every production caller ignored, and which the
///         single-file wrapper discarded along with the batch. A file holding the previous version
///         of the caller's document stayed on disk with nothing said about it anywhere.
///     </para>
/// </summary>
public class CleanupDebtTests : TestBase
{
    [Fact]
    public void AStagingCleanupFailure_ShouldNotReplaceTheFailureThatCausedIt()
    {
        using var batch = new BoundedFileBatch(1_000, "output", Recovery, []);
        batch.BeforeDelete = _ => throw new IOException("the staging file is locked");

        var failure = Assert.Throws<InvalidOperationException>(() =>
            batch.Stage(Path.Combine(TestDir, "unwritable.bin"),
                _ => throw new InvalidOperationException("the renderer failed")));

        // The renderer's failure is the one the caller sees, not the cleanup's.
        Assert.Equal("the renderer failed", failure.Message);
        Assert.Contains(batch.CleanupFailures,
            debt => debt.Contains("the staging file is locked", StringComparison.Ordinal));
    }

    [Fact]
    public void ARefusedStagingPath_ShouldNotLeaveTheDestinationReserved()
    {
        // The staging name was derived outside the try, so a refusal there kept the destination
        // reserved for a file the batch would never produce and no later attempt could use it.
        using var batch = new BoundedFileBatch(1_000, "output", Recovery, []);
        var destination = Path.Combine(TestDir, "retry.bin");

        batch.BeforeDelete = null;
        batch.Stage(destination, stream => stream.Write("first"u8));

        var duplicate = Assert.Throws<ArgumentException>(() =>
            batch.Stage(destination, stream => stream.Write("second"u8)));

        Assert.Contains("already being written", duplicate.Message, StringComparison.Ordinal);

        // The first staging is intact: the refused second attempt released nothing that was not
        // its own.
        batch.Publish();
        Assert.Equal("first", File.ReadAllText(destination));
    }

    [Fact]
    public void ABackupThatCannotBeRemoved_ShouldBeAnnouncedAndReturned()
    {
        var destination = Path.Combine(TestDir, "replaced.bin");
        File.WriteAllText(destination, "previous version");

        var announced = new List<string>();
        long written;
        IReadOnlyList<string> debts;

        using (var batch = new BoundedFileBatch(1_000, "output", Recovery, []))
        {
            batch.ReportDebt = announced.Add;

            // The backup is deleted after the commit point. Failing that delete is what leaves the
            // previous version of the document on disk.
            var deletes = 0;
            batch.BeforeDelete = _ =>
            {
                if (++deletes == 1) throw new IOException("the backup is locked");
            };

            written = batch.Stage(destination, stream => stream.Write("new"u8));
            batch.Publish();
            debts = batch.CleanupFailures;
        }

        Assert.Equal(3, written);
        Assert.Equal("new", File.ReadAllText(destination));

        Assert.NotEmpty(debts);
        Assert.Contains(announced,
            message => message.StartsWith("[WARN]", StringComparison.Ordinal)
                       && message.Contains("still on disk", StringComparison.Ordinal));
    }

    [Fact]
    public void AnOrdinaryPublish_ShouldAnnounceNothing()
    {
        // A guard that reports on every publish would be noise nobody reads.
        var announced = new List<string>();

        using (var batch = new BoundedFileBatch(1_000, "output", Recovery, []))
        {
            batch.ReportDebt = announced.Add;
            batch.Stage(Path.Combine(TestDir, "quiet.bin"),
                stream => stream.Write("ok"u8));
            batch.Publish();
        }

        Assert.Empty(announced);
    }

    [Fact]
    public void TheSingleFilePublisher_ShouldHandBackWhatItCouldNotCleanUp()
    {
        var outcome = BoundedFilePublisher.Publish(Path.Combine(TestDir, "wrapper.bin"), 1_000,
            stream => stream.Write("payload"u8), "output", Recovery, []);

        Assert.Equal(7, outcome.WrittenBytes);
        Assert.Empty(outcome.CleanupFailures);
    }
}
