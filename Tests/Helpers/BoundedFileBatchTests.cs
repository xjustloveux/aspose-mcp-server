using System.Text;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Helpers;

/// <summary>
///     The contract every fan-out handler now shares: a request produces all of its files or none.
///     <para>
///         Each output used to be published the moment it was produced, so a request refused on its
///         last record had already replaced the destinations of every record before it. A caller
///         who retried could not tell which files came from the run that failed (R4-R02, R5-R01).
///     </para>
/// </summary>
public class BoundedFileBatchTests : IDisposable
{
    private readonly string _directory =
        Path.Combine(Path.GetTempPath(), "BoundedFileBatch_" + Guid.NewGuid().ToString("N"));

    /// <summary>Creates the directory the batch writes into.</summary>
    public BoundedFileBatchTests()
    {
        Directory.CreateDirectory(_directory);
    }

    /// <summary>
    ///     This test's recovery context: its own directory, and the key kept there.
    /// </summary>
    /// <remarks>
    ///     Per test class, which is the isolation the production change is about: two classes
    ///     sharing one context would share a journal directory and a signing key (R18-ARCH01).
    /// </remarks>
    private RecoveryContext Recovery => RecoveryContext.For(_directory);

    /// <summary>Removes the directory.</summary>
    public void Dispose()
    {
        if (Directory.Exists(_directory)) Directory.Delete(_directory, true);
        GC.SuppressFinalize(this);
    }

    /// <summary>
    ///     Names a file in this test's directory.
    /// </summary>
    /// <param name="name">Its file name.</param>
    /// <returns>The full path.</returns>
    private string Path_(string name)
    {
        return Path.Combine(_directory, name);
    }

    /// <summary>
    ///     Writes the given bytes to whatever stream the batch supplies.
    /// </summary>
    /// <param name="content">The content to write.</param>
    /// <returns>A writer for the batch.</returns>
    private static Action<Stream> Writes(string content)
    {
        var bytes = Encoding.UTF8.GetBytes(content);
        return stream => stream.Write(bytes, 0, bytes.Length);
    }

    /// <summary>Staging files the batch left behind, if any.</summary>
    /// <returns>Their names.</returns>
    private string[] StagingLeftovers()
    {
        return Directory.GetFiles(_directory)
            .Where(f => Path.GetFileName(f).Contains(".partial-", StringComparison.Ordinal)
                        || Path.GetFileName(f).Contains(".replaced-", StringComparison.Ordinal))
            .ToArray();
    }

    [Fact]
    public void APublishedBatch_ShouldPlaceEveryFileAndLeaveNoStaging()
    {
        using var batch = new BoundedFileBatch(1024, "test files", Recovery);
        batch.Stage(Path_("one.txt"), Writes("first"));
        batch.Stage(Path_("two.txt"), Writes("second"));

        var published = batch.Publish();

        Assert.Equal([Path_("one.txt"), Path_("two.txt")], published);
        Assert.Equal("first", File.ReadAllText(Path_("one.txt")));
        Assert.Equal("second", File.ReadAllText(Path_("two.txt")));
        Assert.Empty(StagingLeftovers());
    }

    /// <summary>
    ///     The case the old contract got wrong: the budget runs out on the last file. The earlier
    ///     ones were already at their destinations, replacing whatever was there.
    /// </summary>
    [Fact]
    public void ABatchRefusedOnItsLastFile_ShouldLeaveEveryDestinationAsItWas()
    {
        File.WriteAllText(Path_("one.txt"), "the caller's own file");
        File.WriteAllText(Path_("two.txt"), "another of the caller's files");

        var refused = Assert.Throws<ArgumentException>(() =>
        {
            using var batch = new BoundedFileBatch(12, "test files", Recovery);
            batch.Stage(Path_("one.txt"), Writes("0123456789"));
            batch.Stage(Path_("two.txt"), Writes("this one passes the budget"));
        });

        Assert.Contains("test files", refused.Message, StringComparison.Ordinal);
        Assert.Equal("the caller's own file", File.ReadAllText(Path_("one.txt")));
        Assert.Equal("another of the caller's files", File.ReadAllText(Path_("two.txt")));
        Assert.Empty(StagingLeftovers());
    }

    /// <summary>
    ///     A budget already spent refuses before writing anything, rather than producing a file and
    ///     measuring it afterwards.
    /// </summary>
    [Fact]
    public void ABatchWithNothingLeft_ShouldRefuseBeforeWriting()
    {
        using (var batch = new BoundedFileBatch(5, "test files", Recovery))
        {
            batch.Stage(Path_("one.txt"), Writes("12345"));

            // Refused before anything is written, rather than written and measured afterwards.
            Assert.Throws<ArgumentException>(() => batch.Stage(Path_("two.txt"), Writes("x")));
            Assert.Equal(1, batch.Count);
        }

        Assert.False(File.Exists(Path_("one.txt")));
        Assert.False(File.Exists(Path_("two.txt")));
        Assert.Empty(StagingLeftovers());
    }

    /// <summary>
    ///     A producer that throws part-way — a serializer refusing the content it was handed — is
    ///     the same failure as an exhausted budget as far as the destinations are concerned.
    /// </summary>
    [Fact]
    public void AProducerThatThrows_ShouldLeaveEveryDestinationAsItWas()
    {
        File.WriteAllText(Path_("one.txt"), "the caller's own file");

        Assert.Throws<InvalidOperationException>(() =>
        {
            using var batch = new BoundedFileBatch(1024, "test files", Recovery);
            batch.Stage(Path_("one.txt"), Writes("replacement"));
            batch.Stage(Path_("two.txt"), _ => throw new InvalidOperationException("serializer gave up"));
        });

        Assert.Equal("the caller's own file", File.ReadAllText(Path_("one.txt")));
        Assert.False(File.Exists(Path_("two.txt")));
        Assert.Empty(StagingLeftovers());
    }

    [Fact]
    public void APublishedBatch_ShouldReplaceAnExistingDestination()
    {
        File.WriteAllText(Path_("one.txt"), "the previous run's output");

        using (var batch = new BoundedFileBatch(1024, "test files", Recovery))
        {
            batch.Stage(Path_("one.txt"), Writes("this run's output"));
            batch.Publish();
        }

        Assert.Equal("this run's output", File.ReadAllText(Path_("one.txt")));
        Assert.Empty(StagingLeftovers());
    }

    /// <summary>
    ///     §17.3.3: the window between taking a backup and putting the staged file in its place.
    ///     <para>
    ///         The pair was recorded only after <em>both</em> moves succeeded, so a stage move that
    ///         failed after its backup had been taken was not in the rollback journal at all: the
    ///         caller's file stayed in the backup and the destination was simply gone. Every step is
    ///         written down before it is taken now.
    ///     </para>
    /// </summary>
    [Fact]
    public void APublishThatFailsAfterTakingABackup_ShouldPutTheDestinationBack()
    {
        File.WriteAllText(Path_("one.txt"), "first destination");
        File.WriteAllText(Path_("two.txt"), "the caller's own file");

        using (var batch = new BoundedFileBatch(1024, "test files", Recovery))
        {
            batch.Stage(Path_("one.txt"), Writes("first replacement"));
            batch.Stage(Path_("two.txt"), Writes("second replacement"));

            // Fail exactly the second entry's stage move — after its backup has been taken.
            batch.BeforeMove = (source, destination) =>
            {
                if (source.Contains(".partial-", StringComparison.Ordinal)
                    && destination.EndsWith("two.txt", StringComparison.Ordinal))
                    throw new IOException("the stage move failed");
            };

            Assert.Throws<IOException>(() => batch.Publish());

            // The destination that was already backed up is back, and so is the one published
            // before it. Its own staging file is still the batch's to clean up on the way out.
            Assert.Equal("the caller's own file", File.ReadAllText(Path_("two.txt")));
            Assert.Equal("first destination", File.ReadAllText(Path_("one.txt")));
            Assert.Empty(batch.RollbackFailures);
        }

        Assert.Empty(StagingLeftovers());
    }

    /// <summary>
    ///     §17.3.3: a rollback that cannot put a destination back is reported rather than hidden
    ///     behind the original failure.
    /// </summary>
    [Fact]
    public void ARollbackThatCannotRestore_ShouldReportItAlongsideTheOriginalFailure()
    {
        File.WriteAllText(Path_("one.txt"), "the caller's own file");

        using var batch = new BoundedFileBatch(1024, "test files", Recovery);
        batch.Stage(Path_("one.txt"), Writes("replacement"));

        batch.BeforeMove = (source, destination) =>
        {
            if (source.Contains(".partial-", StringComparison.Ordinal))
                throw new IOException("the stage move failed");
            if (source.Contains(".replaced-", StringComparison.Ordinal))
                throw new IOException("and the restore failed too");
            _ = destination;
        };

        var failure = Assert.Throws<AggregateException>(() => batch.Publish());

        Assert.Contains(failure.InnerExceptions, e => e.Message.Contains("the stage move failed"));
        Assert.Contains(failure.InnerExceptions, e => e.Message.Contains("could not be restored"));
        Assert.NotEmpty(batch.RollbackFailures);
    }

    /// <summary>
    ///     §17.3.4: the same destination twice would make the publish order pick a winner and the
    ///     rollback restore the wrong file.
    /// </summary>
    [Fact]
    public void ABatchStagingOneDestinationTwice_ShouldBeRefused()
    {
        using var batch = new BoundedFileBatch(1024, "test files", Recovery);
        batch.Stage(Path_("one.txt"), Writes("first"));

        var refusal = Assert.Throws<ArgumentException>(() => batch.Stage(Path_("one.txt"), Writes("second")));

        Assert.Contains("already being written", refusal.Message, StringComparison.Ordinal);
        Assert.Equal(1, batch.Count);
    }

    /// <summary>
    ///     §17.3.4: a destination validated once says nothing about the derived staging and backup
    ///     paths, so every path the batch touches is re-checked against the allowlist as it acts.
    /// </summary>
    [Fact]
    public void ABatchWithAnAllowlist_ShouldRefuseADestinationOutsideIt()
    {
        var inside = Path.Combine(_directory, "inside");
        Directory.CreateDirectory(inside);

        using var batch = new BoundedFileBatch(1024, "test files", Recovery, [inside]);

        batch.Stage(Path.Combine(inside, "ok.txt"), Writes("fine"));
        Assert.Throws<ArgumentException>(() => batch.Stage(Path_("outside.txt"), Writes("no")));
        Assert.False(File.Exists(Path_("outside.txt")));
    }

    /// <summary>
    ///     §17.3.4: a parent directory replaced by a symlink out of the allowlist between staging
    ///     and publishing must be caught when the move happens, not only when the path was first
    ///     accepted.
    /// </summary>
    [SkippableFact]
    public void APublishWhoseParentBecomesALinkOutOfTheAllowlist_ShouldBeRefused()
    {
        var inside = Path.Combine(_directory, "inside");
        var elsewhere = Path.Combine(_directory, "elsewhere");
        Directory.CreateDirectory(inside);
        Directory.CreateDirectory(elsewhere);

        using var batch = new BoundedFileBatch(1024, "test files", Recovery, [inside]);
        var destination = Path.Combine(inside, "one.txt");
        batch.Stage(destination, Writes("staged"));

        // Between staging and publishing, the destination becomes a link out of the allowlist.
        Skip.IfNot(SymlinkFixture.TryCreateFileSymlink(destination,
                Path.Combine(elsewhere, "one.txt")),
            "This platform or account cannot create symbolic links.");

        Assert.Throws<ArgumentException>(() => batch.Publish());
        Assert.False(File.Exists(Path.Combine(elsewhere, "one.txt")),
            "the publish followed the link out of the allowlist");
    }

    /// <summary>
    ///     §18.4.2: once every output is in place the request has succeeded. Removing the backups
    ///     is tidying up afterwards, and a failure there used to be reported as a failed publish —
    ///     with the new outputs already at their destinations, which is the one thing a handled
    ///     failure promises not to do.
    /// </summary>
    [Fact]
    public void ACleanupFailureAfterTheCommit_ShouldBeAWarningRatherThanAFailure()
    {
        File.WriteAllText(Path_("one.txt"), "the previous run's output");

        using var batch = new BoundedFileBatch(1024, "test files", Recovery);
        batch.Stage(Path_("one.txt"), Writes("this run's output"));

        // The one failure a filesystem will not produce on demand: the backup cannot be removed
        // after every output is already in place.
        batch.BeforeDelete = path =>
        {
            if (path.Contains(".replaced-", StringComparison.Ordinal))
                throw new IOException("the backup could not be removed");
        };

        var published = batch.Publish();

        // The request succeeded: the output is at its destination and the caller is told what was
        // left behind rather than being told the publish failed.
        Assert.Single(published);
        Assert.Equal("this run's output", File.ReadAllText(Path_("one.txt")));
        Assert.NotEmpty(batch.CleanupFailures);
        Assert.Contains(batch.CleanupFailures,
            message => message.Contains("could not be removed", StringComparison.Ordinal));
    }

    /// <summary>
    ///     §18.4.2: a batch is one transaction, so nothing may be staged into it once it has
    ///     published. Such a file would never be published and Dispose would not remove it either.
    /// </summary>
    [Fact]
    public void ABatchThatHasPublished_ShouldRefuseToStageMore()
    {
        using var batch = new BoundedFileBatch(1024, "test files", Recovery);
        batch.Stage(Path_("one.txt"), Writes("first"));
        batch.Publish();

        var refusal = Assert.Throws<ArgumentException>(() => batch.Stage(Path_("two.txt"), Writes("second")));

        Assert.Contains("already published", refusal.Message, StringComparison.Ordinal);
        Assert.False(File.Exists(Path_("two.txt")));
        Assert.Empty(StagingLeftovers());
    }

    /// <summary>
    ///     §18.4.2: two names differing only in case are one file on Windows and macOS but two on
    ///     Linux, and comparing case-insensitively everywhere refused a legitimate pair there.
    /// </summary>
    [Fact]
    public void ABatchStagingNamesThatDifferOnlyInCase_ShouldFollowTheFilesystem()
    {
        using var batch = new BoundedFileBatch(1024, "test files", Recovery);
        batch.Stage(Path_("case.txt"), Writes("first"));

        if (OperatingSystem.IsWindows() || OperatingSystem.IsMacOS())
        {
            Assert.Throws<ArgumentException>(() => batch.Stage(Path_("CASE.txt"), Writes("second")));
            Assert.Equal(1, batch.Count);
        }
        else
        {
            batch.Stage(Path_("CASE.txt"), Writes("second"));
            Assert.Equal(2, batch.Count);
        }
    }

    [Fact]
    public void ABatchThatWasNeverPublished_ShouldRemoveWhatItStaged()
    {
        using (var batch = new BoundedFileBatch(1024, "test files", Recovery))
        {
            batch.Stage(Path_("one.txt"), Writes("never published"));
            Assert.Equal(1, batch.Count);
        }

        Assert.False(File.Exists(Path_("one.txt")));
        Assert.Empty(StagingLeftovers());
        Assert.Empty(Directory.GetFiles(_directory));
    }

    [SkippableFact]
    public void TwoSpellingsOfOneWindowsAllowlist_ShouldBeOneHost()
    {
        // R21-REC07. Sorted case-insensitively, compared case-sensitively.
        Skip.IfNot(OperatingSystem.IsWindows(), "Casing is a Windows question.");

        var upper = Path.Combine(_directory, "Roots");
        var lower = Path.Combine(_directory.ToLowerInvariant(), "roots");

        Assert.Equal(BoundedFileBatch.SinkKey([upper]), BoundedFileBatch.SinkKey([lower]));
    }

    [Fact]
    public void OnePathContainingTheSeparator_ShouldNotBeTwoPaths()
    {
        // R22-REC03. The key is a serialisation, and a serialisation joined on a character a
        // path may contain is ambiguous: one path with a newline in it read as two. The key
        // never touches the filesystem, so the shape is the same on every platform.
        var first = Path.Combine(_directory, "a");
        var second = Path.Combine(_directory, "b");

        Assert.NotEqual(BoundedFileBatch.SinkKey([first + "\n" + second]),
            BoundedFileBatch.SinkKey([first, second]));
    }
}
