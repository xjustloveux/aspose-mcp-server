using System.Security.Cryptography;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Helpers;

/// <summary>
///     §23.19.2: what a multi-file publish leaves behind when the process dies in the middle of it.
///     <para>
///         <see cref="BoundedFileBatch" /> rolls back within the process — a move that fails
///         restores every destination it had replaced. A crash is different: nothing runs the
///         rollback, and the tree is left with some destinations replaced, their previous files
///         sitting as <c>.replaced-*</c> siblings, and no record of which was which.
///     </para>
///     <para>
///         A crash cannot be staged, so the journal is what stands in for one: it exists on disk
///         only while a publish is in flight, so a test can write the state a crash would leave and
///         ask recovery to put it right. That is the same thing recovery will face at startup.
///     </para>
/// </summary>
[Collection("SerialStaticSeams")]
public class PublishJournalTests : TestBase
{
    /// <summary>This test's signing key, kept in its own directory.</summary>
    /// <returns>The capability every fixture signs and verifies with.</returns>
    private RecoveryCapability Capability()
    {
        return Recovery.Capability ?? throw new InvalidOperationException(
            "the test directory could not hold a recovery key");
    }

    /// <summary>Writes a record the way a real publish would: signed, and still publishing.</summary>
    /// <param name="path">Where the journal goes.</param>
    /// <param name="what">What the batch was publishing.</param>
    /// <param name="steps">Its steps.</param>
    /// <param name="state">The transaction's state.</param>
    private void WriteSigned(string path, string what, PublishJournal.Entry[] steps,
        string state = PublishJournal.Publishing)
    {
        var record = PublishJournal.Signed(
            new PublishJournal.Record(what, steps, State: state), Capability());
        if (Path.GetFileName(path).EndsWith(PublishJournal.Extension, StringComparison.Ordinal))
            PublishJournal.WriteManaged(path, record);
        else
            PublishJournal.Write(path, record);
    }

    /// <summary>Builds the state a crash mid-publish leaves: destination replaced, backup beside it.</summary>
    /// <param name="name">Base file name.</param>
    /// <returns>The destination, its backup, and the journal describing them.</returns>
    private (string Destination, string Backup, string Journal) AnInterruptedPublish(string name)
    {
        var destination = CreateTestFilePath(name);
        var backup = destination + ".replaced-" + Guid.NewGuid().ToString("N");

        File.WriteAllText(backup, "the previous version");
        File.WriteAllText(destination, "the half-written new version");

        var journal = CreateTestFilePath(name + PublishJournal.Extension);
        WriteSigned(journal, "a fixture",
            [new PublishJournal.Entry(destination, backup)]);

        return (destination, backup, journal);
    }

    [Fact]
    public void AnInterruptedPublish_ShouldBePutBackByRecovery()
    {
        var (destination, backup, journal) = AnInterruptedPublish("journal_restore.txt");

        var result = PublishJournal.Recover(journal, [TestDir], Capability());

        Assert.Contains(destination, result.Restored);
        Assert.Empty(result.Refused);
        Assert.Equal("the previous version", File.ReadAllText(destination));
        Assert.False(File.Exists(backup), "the backup should have moved back, not been copied");
        Assert.False(File.Exists(journal), "a completed recovery leaves no journal");
    }

    [Fact]
    public void AJournalWhoseBackupIsGone_ShouldBeTreatedAsAFinishedPublish()
    {
        // The publish reached its commit point and tidied the backup; the journal is simply stale.
        // Restoring anything here would undo a request that succeeded.
        var (destination, backup, journal) = AnInterruptedPublish("journal_finished.txt");
        File.Delete(backup);
        File.WriteAllText(destination, "the new version, complete");

        var result = PublishJournal.Recover(journal, [TestDir], Capability());

        Assert.Empty(result.Restored);
        Assert.Equal("the new version, complete", File.ReadAllText(destination));
        Assert.False(File.Exists(journal));
    }

    [Fact]
    public void AJournalPointingOutsideTheTrustedRoots_ShouldBeRefused()
    {
        // The roots come from the caller now. The journal used to carry its own, and recovery
        // read them back — so a file dropped into a scanned directory chose the roots it wanted
        // (R13-SEC01). Here the caller trusts somewhere else entirely, and the entry is refused.
        var destination = CreateTestFilePath("journal_outside.txt");
        var backup = destination + ".replaced-" + new string('a', 32);
        File.WriteAllText(backup, "previous");
        File.WriteAllText(destination, "current");

        var journal = CreateTestFilePath("journal_outside.json");
        WriteSigned(journal, "a fixture",
            [new PublishJournal.Entry(destination, backup)]);

        var result = PublishJournal.Recover(journal, [Path.Combine(TestDir, "somewhere_else")], Capability());

        Assert.Empty(result.Restored);
        Assert.Single(result.Refused);
        Assert.Equal("current", File.ReadAllText(destination));
        Assert.True(File.Exists(journal), "a journal it could not act on is kept, not deleted");
    }

    [Fact]
    public void RecoveryWithNoTrustedRoots_ShouldActOnASignedRecord()
    {
        // No allowlist is the default: the server may write anywhere the caller names, and that is
        // where its outputs are. This used to refuse — "no roots means no authority" — which was
        // right while the paths were the only thing recovery had to judge by, and left the default
        // configuration with a crash recovery that recovered nothing, in exactly the situation it
        // existed for (R17-F02). The signature is the authority now.
        var destination = CreateTestFilePath("journal_noroots.txt");
        var backup = destination + ".replaced-" + new string('b', 32);
        File.WriteAllText(backup, "the previous version");
        File.WriteAllText(destination, "the half-written new version");

        var journal = CreateTestFilePath("journal_noroots.json");
        WriteSigned(journal, "a fixture", [new PublishJournal.Entry(destination, backup)]);

        var result = PublishJournal.Recover(journal, [], Capability());

        Assert.Contains(destination, result.Restored);
        Assert.Equal("the previous version", File.ReadAllText(destination));
    }

    [Fact]
    public void RecoveryWithNoTrustedRoots_ShouldStillRefuseARecordItDidNotWrite()
    {
        // The half that has to survive relaxing the roots. Without an allowlist there is no
        // containment left, so the signature is the only thing standing between a file someone
        // dropped in this directory and a filesystem operation performed as the server.
        var destination = CreateTestFilePath("journal_unsigned.txt");
        var backup = destination + ".replaced-" + new string('f', 32);
        File.WriteAllText(backup, "the attacker's bytes");
        File.WriteAllText(destination, "a document this server produced earlier");

        var journal = CreateTestFilePath("journal_unsigned.json");
        PublishJournal.Write(journal, new PublishJournal.Record("a forged record",
            [new PublishJournal.Entry(destination, backup)]));

        var result = PublishJournal.Recover(journal, [], Capability());

        Assert.Empty(result.Restored);
        Assert.NotEmpty(result.Refused);
        Assert.Equal("a document this server produced earlier", File.ReadAllText(destination));
        Assert.True(File.Exists(journal),
            "an unsigned journal is left alone, not deleted — removing it would let anyone who can "
            + "write here erase a real interrupted publish's record");
    }

    [Fact]
    public void ARecordSignedByADifferentInstallation_ShouldBeRefused()
    {
        // A key from somewhere else is not this installation's key, and a record carrying a
        // well-formed signature that does not verify is exactly what a forgery looks like.
        var elsewhere = Path.Combine(TestDir, "another_installation");
        Directory.CreateDirectory(elsewhere);
        var theirs = RecoveryCapability.For(elsewhere);
        Assert.NotNull(theirs);

        var destination = CreateTestFilePath("journal_foreign.txt");
        var backup = destination + ".replaced-" + new string('9', 32);
        File.WriteAllText(backup, "their bytes");
        File.WriteAllText(destination, "ours");

        var journal = CreateTestFilePath("journal_foreign.json");
        PublishJournal.Write(journal, PublishJournal.Signed(
            new PublishJournal.Record("signed elsewhere",
                [new PublishJournal.Entry(destination, backup)],
                State: PublishJournal.Publishing),
            theirs));

        var result = PublishJournal.Recover(journal, [TestDir], Capability());

        Assert.Empty(result.Restored);
        Assert.Equal("ours", File.ReadAllText(destination));
    }

    [Fact]
    public void ACommittedTransaction_ShouldNeverHaveItsDestinationsTouched()
    {
        // The publish succeeded; only its journal outlived it, because deleting that journal
        // failed and the failure was swallowed. This used to be read as an interrupted publish —
        // and for a created step the recorded digest matches the delivered file *by construction*,
        // so the rollback deleted exactly what had just been published (R17-F01).
        var replaced = CreateTestFilePath("committed_replaced.txt");
        var backup = replaced + ".replaced-" + new string('7', 32);
        File.WriteAllText(backup, "the version before this publish");
        File.WriteAllText(replaced, "the delivered version");

        var created = CreateTestFilePath("committed_created.txt");
        File.WriteAllText(created, "a file this publish created");

        var journal = CreateTestFilePath("committed.json");
        WriteSigned(journal, "a finished publish",
            [
                new PublishJournal.Entry(replaced, backup),
                new PublishJournal.Entry(created, null, PublishJournal.DigestOf(created))
            ],
            PublishJournal.Committed);

        var result = PublishJournal.Recover(journal, [TestDir], Capability());

        Assert.Empty(result.Restored);
        Assert.Equal("the delivered version", File.ReadAllText(replaced));
        Assert.True(File.Exists(created), "a committed publish's output was deleted");
        Assert.Equal("a file this publish created", File.ReadAllText(created));

        // And the leftovers it did have are gone.
        Assert.False(File.Exists(backup), "the superseded backup was left behind");
        Assert.False(File.Exists(journal), "the committed record was left behind");
    }

    [Fact]
    public void AnEntryWhoseBackupIsNotTheSiblingThisCodeWrites_ShouldBeRefused()
    {
        // Both halves of the move are attacker-chosen otherwise. A backup this code produced is
        // the destination's own name plus `.replaced-` and a 32-hex nonce; anything else names a
        // pair no publish here ever created (R13-SEC01).
        var destination = CreateTestFilePath("journal_shape.txt");
        var elsewhere = CreateTestFilePath("unrelated_source.txt");
        File.WriteAllText(elsewhere, "not a backup of anything");
        File.WriteAllText(destination, "current");

        var journal = CreateTestFilePath("journal_shape.json");
        WriteSigned(journal, "a fixture",
            [new PublishJournal.Entry(destination, elsewhere)]);

        var result = PublishJournal.Recover(journal, [TestDir], Capability());

        Assert.Empty(result.Restored);
        Assert.NotEmpty(result.Refused);
        Assert.Equal("current", File.ReadAllText(destination));
        Assert.True(File.Exists(elsewhere), "the unrelated file was moved");
    }

    [Fact]
    public void AUncOrDevicePath_ShouldBeRefusedBeforeTheFilesystemIsAsked()
    {
        // The lookup is the effect: `File.Exists` on a crafted UNC path is a network request the
        // journal caused. Validation is structural and runs first, so nothing is asked (R13-SEC01).
        var asked = new List<string>();

        var destination = @"\\attacker.invalid\share\payload.txt";
        var journal = CreateTestFilePath("journal_unc.json");
        WriteSigned(journal, "a fixture",
            [new PublishJournal.Entry(destination, destination + ".replaced-" + new string('c', 32))]);

        var result = PublishJournal.Recover(journal, [TestDir], Capability(),
            (from, to) => asked.Add($"{from} -> {to}"));

        Assert.Empty(result.Restored);
        Assert.Empty(asked);
        Assert.NotEmpty(result.Refused);
    }

    [Fact]
    public void AnInterruptedPublishThatOnlyCreatedFiles_ShouldReportOnesItCannotIdentify()
    {
        // It used to delete them, and that was the one thing recovery must not do on the strength
        // of a journal. Nothing binds a journal to a publish this server started, so a file named
        // `<anything>.publishing.json` holding `{"Steps":[{"Destination":"<any file>"}]}` was an
        // arbitrary-delete primitive that ran at every start — a capability a caller who could
        // write those files did not otherwise have. Restoring a backup gives them nothing they
        // did not already have, so that half stays; deleting does, so it went (independent
        // review of R13-SEC01/R13-F01).
        var created = CreateTestFilePath("journal_created.txt");
        File.WriteAllText(created, "half a transaction");

        var journal = CreateTestFilePath("journal_created.json");

        // No digest: the entry names a file and says nothing about what is in it.
        WriteSigned(journal, "a fixture",
            [new PublishJournal.Entry(created, null)]);

        var result = PublishJournal.Recover(journal, [TestDir], Capability());

        Assert.Contains(created, result.Orphaned);
        Assert.Empty(result.Restored);
        Assert.True(File.Exists(created), "recovery deleted a file named by a journal");
        Assert.Equal("half a transaction", File.ReadAllText(created));
        Assert.False(File.Exists(journal), "the journal was kept when nothing was refused");
    }

    [Fact]
    public void APublishIntoADirectoryOutsideTheRecoveryOne_ShouldStillLeaveAJournalWhereRecoveryLooks()
    {
        // The default configuration: no allowlist, so the caller names the output directory and it
        // is not the server's temp directory. The batch used to write its journal beside the
        // output while recovery only ever scanned the temp directory and the allowlist — so a real
        // crash, in the shape most deployments run, left a real journal somewhere nothing would
        // ever look at again. The same defect R13-F01 was supposed to have closed, one layer down
        // (independent review).
        var host = Path.Combine(TestDir, "server_temp");
        var output = Path.Combine(TestDir, "somewhere_the_caller_named");
        Directory.CreateDirectory(host);
        Directory.CreateDirectory(output);

        // A context of this host's own, which is what a batch is handed now — the journal goes
        // where that host recovers from, not beside the output (R18-ARCH01).
        var hostRecovery = RecoveryContext.For(host);
        var recovery = hostRecovery.Directory;

        var destination = Path.Combine(output, "published.txt");
        File.WriteAllText(destination, "the previous version");

        string[] duringTheMove = [];

        using (var batch = new BoundedFileBatch(4096, "a fixture", hostRecovery, []))
        {
            batch.ReportDebt = _ => { };
            batch.Stage(destination, stream => stream.Write("the new version"u8));
            batch.BeforeMove = (_, _) =>
            {
                if (duringTheMove.Length == 0)
                    duringTheMove = Directory.GetFiles(recovery, "*" + PublishJournal.Extension);
            };

            batch.Publish();
        }

        Assert.NotEmpty(duringTheMove);
    }

    [Fact]
    public void AJournalNamingAFileItNeverCreated_ShouldNotBeAbleToDeleteIt()
    {
        // The exploit written out. The journal is entirely attacker-chosen and names a file that
        // has nothing to do with any publish; every other rule passes, because the path is
        // canonical, under the trusted root, and not a link.
        var victim = CreateTestFilePath("someone_elses_document.docx");
        File.WriteAllText(victim, "a document this server produced earlier");

        var journal = CreateTestFilePath("planted" + PublishJournal.Extension);
        WriteSigned(journal, "a forged record",
            [new PublishJournal.Entry(victim, null)]);

        // Naming the file is all a forger who cannot read it can do, and naming it is no longer
        // enough.

        PublishJournal.Recover(journal, [TestDir], Capability());

        Assert.True(File.Exists(victim), "a planted journal deleted an unrelated file");
        Assert.Equal("a document this server produced earlier", File.ReadAllText(victim));
    }

    [Fact]
    public void ACreatedFileThatIsStillWhatThePublishWrote_ShouldBeRolledBack()
    {
        // The feature the delete existed for, given back: a crash partway through a batch of new
        // files leaves them behind, and the next start clears them. Safe because the journal says
        // what the file contains, not merely where it is.
        var created = CreateTestFilePath("journal_created_matching.txt");
        File.WriteAllText(created, "half a transaction");

        var journal = CreateTestFilePath("journal_created_matching.json");
        WriteSigned(journal, "a fixture",
            [new PublishJournal.Entry(created, null, PublishJournal.DigestOf(created))]);

        var result = PublishJournal.Recover(journal, [TestDir], Capability());

        Assert.Contains(created, result.Restored);
        Assert.False(File.Exists(created), "the file this transaction created is still there");
    }

    [Fact]
    public void ACreatedFileWhoseContentsChanged_ShouldBeLeftAlone()
    {
        // The forgery, and also the honest race: whatever is at that path is not what the publish
        // wrote, so removing it would be destroying something else. A journal naming a victim file
        // and a digest that does not match it — which is every digest an attacker can produce for
        // a file they cannot read — takes this path.
        var victim = CreateTestFilePath("someone_elses_report.docx");
        File.WriteAllText(victim, "a document this server produced earlier");

        var journal = CreateTestFilePath("journal_created_mismatch.json");
        WriteSigned(journal, "a forged record",
            [new PublishJournal.Entry(victim, null, new string('0', 64))]);

        var result = PublishJournal.Recover(journal, [TestDir], Capability());

        Assert.Contains(victim, result.Orphaned);
        Assert.Empty(result.Restored);
        Assert.True(File.Exists(victim), "a journal with a wrong digest still deleted the file");
        Assert.Equal("a document this server produced earlier", File.ReadAllText(victim));
    }

    [Fact]
    public void ABatchThatCreatesFiles_ShouldRecordWhatEachOneWillContain()
    {
        // Driven through the batch, so this is about what production records rather than about
        // what a hand-written journal can express. Without the digest the entry cannot be told
        // apart from one naming any other file.
        var first = CreateTestFilePath("digest_one.txt");
        var second = CreateTestFilePath("digest_two.txt");

        Assert.False(File.Exists(first), "the fixture must be creating files, not replacing them");

        string? recorded = null;

        using (var batch = new BoundedFileBatch(4096, "a fixture", Recovery, [TestDir]))
        {
            batch.Stage(first, stream => stream.Write("one"u8));
            batch.Stage(second, stream => stream.Write("two"u8));
            batch.BeforeMove = (_, _) =>
            {
                var found = Directory.EnumerateFiles(Recovery.Directory, "*" + PublishJournal.Extension)
                    .FirstOrDefault();
                if (found != null) recorded ??= File.ReadAllText(found);
            };

            batch.Publish();
        }

        Assert.NotNull(recorded);

        // The digests of "one" and "two", which is what those two files end up holding.
        Assert.Contains(
            Convert.ToHexString(SHA256.HashData("one"u8))
                .ToLowerInvariant(),
            recorded!, StringComparison.Ordinal);
        Assert.Contains(
            Convert.ToHexString(SHA256.HashData("two"u8))
                .ToLowerInvariant(),
            recorded!, StringComparison.Ordinal);
    }

    [Fact]
    public void AJournalWhoseWriterIsStillRunning_ShouldBeLeftAlone()
    {
        // A journal exists for as long as its publish runs, so one written by a live process is a
        // transaction in flight, not a crash. Two hosts sharing a temp directory is a shape the
        // debt sink stack already anticipates, and rolling back the other one's publish mid-way
        // would corrupt exactly what this protects.
        var destination = CreateTestFilePath("journal_live.txt");
        var backup = destination + ".replaced-" + new string('d', 32);
        File.WriteAllText(backup, "the previous version");
        File.WriteAllText(destination, "the half-written new version");

        var journal = CreateTestFilePath("journal_live.json");
        PublishJournal.Write(journal, PublishJournal.Describe("a publish in flight",
            [new PublishJournal.Entry(destination, backup)], PublishJournal.Publishing));

        var result = PublishJournal.Recover(journal, [TestDir], Capability());

        Assert.Empty(result.Restored);
        Assert.Equal("the half-written new version", File.ReadAllText(destination));
        Assert.True(File.Exists(journal), "a live publish's journal was deleted");
    }

    [SkippableFact]
    public void ADestinationReachedThroughAJunction_ShouldBeRefused()
    {
        // Two defects in one check. It read `Backup ?? Destination`, so in the restore branch the
        // file being written *over* — the one that decides where the bytes land — was never
        // examined (independent review). And it asked only about the leaf, so a destination whose
        // *parent* is a junction pointing out of the trusted root passed: the path is lexically
        // canonical, lexically under the root, and the file at the end of it is perfectly
        // ordinary. That is the containment failure the cleanup queue fixed in §21.2.2 and the
        // journal never inherited; both now call the same walk.
        //
        // A junction rather than a symbolic link, because a junction needs no elevation — which is
        // also why the symbolic-link version of this fixture could only ever skip.
        var inside = Path.Combine(TestDir, "inside_the_root");
        var outside = Path.Combine(TestDir, "outside_the_root");
        Directory.CreateDirectory(outside);

        Skip.IfNot(MidChainLinkFixture.TryCreateDirectoryLink(inside, outside),
            "This host does not allow creating a directory link");

        var destination = Path.Combine(inside, "victim.txt");
        File.WriteAllText(destination, "what the journal never approved touching");

        var backup = destination + ".replaced-" + new string('e', 32);
        File.WriteAllText(backup, "the payload");

        var journal = CreateTestFilePath("journal_link.json");
        WriteSigned(journal, "a fixture",
            [new PublishJournal.Entry(destination, backup)]);

        var result = PublishJournal.Recover(journal, [TestDir], Capability());

        Assert.Empty(result.Restored);
        Assert.NotEmpty(result.Refused);
        Assert.Equal("what the journal never approved touching", File.ReadAllText(destination));
    }

    [Fact]
    public void MoreJournalsThanOneScanWillTake_ShouldLeaveTheRestForLater()
    {
        // A cap on each file's size says nothing about how many files there are, and the
        // directory is one anything that can write output can add to.
        for (var i = 0; i < 600; i++)
            PublishJournal.Write(CreateTestFilePath($"flood_{i}" + PublishJournal.Extension),
                new PublishJournal.Record("a fixture",
                    [new PublishJournal.Entry(CreateTestFilePath($"flood_{i}.txt"), null)]));

        var result = PublishJournal.RecoverAll(TestDir, [TestDir], Capability());

        Assert.Contains(result.Refused, refusal =>
            refusal.Contains("remaining budget", StringComparison.Ordinal));
    }

    [Fact]
    public void APublishCreatingNewFiles_ShouldHaveJournalledThemBeforeTheFirstMove()
    {
        // Driven through BoundedFileBatch, not by writing a journal here. A fixture that writes
        // its own journal proves recovery works and says nothing about whether anything records
        // a publish — removing the production write left such a fixture green (R13-F01).
        var first = CreateTestFilePath("created_one.txt");
        var second = CreateTestFilePath("created_two.txt");

        Assert.False(File.Exists(first), "the fixture must be publishing new files, not replacing");

        string? journalAtFirstMove = null;

        using (var batch = new BoundedFileBatch(4096, "a fixture", Recovery, [TestDir]))
        {
            batch.Stage(first, stream => stream.Write("one"u8));
            batch.Stage(second, stream => stream.Write("two"u8));

            // The instant a crash would interrupt: the first move is about to happen.
            batch.BeforeMove = (_, _) =>
            {
                // Read now, not afterwards: a successful publish deletes its journal, so by the
                // time the batch is disposed there is nothing left to look at.
                var found = Directory
                    .EnumerateFiles(Recovery.Directory, "*" + PublishJournal.Extension)
                    .FirstOrDefault();

                if (found != null) journalAtFirstMove ??= File.ReadAllText(found);
            };

            batch.Publish();
        }

        Assert.NotNull(journalAtFirstMove);

        // And what it recorded has to name the created destinations, or recovery has nothing to
        // undo them with.
        Assert.Contains(Path.GetFileName(first), journalAtFirstMove!, StringComparison.Ordinal);
        Assert.Contains(Path.GetFileName(second), journalAtFirstMove!, StringComparison.Ordinal);

        Assert.Empty(Directory.EnumerateFiles(Recovery.Directory, "*" + PublishJournal.Extension));
    }

    [Fact]
    public void AJournalLargerThanAJournalMayBe_ShouldBeRefusedWithoutBeingParsed()
    {
        // Its size is an attacker's choice: the file sits in a directory this server scans at
        // startup.
        var journal = CreateTestFilePath("journal_huge.json");
        File.WriteAllText(journal, new string('x', 5 * 1024 * 1024));

        var result = PublishJournal.Recover(journal, [TestDir], Capability());

        Assert.Empty(result.Restored);
        Assert.NotEmpty(result.Refused);
        Assert.True(File.Exists(journal), "an oversized journal is kept, not deleted");
    }

    [Fact]
    public void AnUnreadableJournal_ShouldBeKeptAndReported()
    {
        // Deleting it would destroy the only record that a publish was interrupted.
        var journal = CreateTestFilePath("journal_corrupt.publishing.json");
        File.WriteAllText(journal, "{ not json");

        var result = PublishJournal.Recover(journal, [TestDir], Capability());

        Assert.Empty(result.Restored);
        Assert.Single(result.Refused);
        Assert.True(File.Exists(journal));
    }

    [Fact]
    public void RecoveryThatFails_ShouldKeepTheJournalForTheNextStart()
    {
        var (destination, _, journal) = AnInterruptedPublish("journal_retry.txt");

        var result = PublishJournal.Recover(journal, [TestDir], Capability(),
            (_, _) => throw new IOException("the destination is locked"));

        Assert.Empty(result.Restored);
        Assert.Single(result.Refused);
        Assert.True(File.Exists(journal), "an unfinished recovery must be retried, not forgotten");
        Assert.Equal("the half-written new version", File.ReadAllText(destination));
    }

    [Fact]
    public void RecoverAll_ShouldFindEveryJournalInTheDirectory()
    {
        var first = AnInterruptedPublish("journal_all_a.txt");
        var second = AnInterruptedPublish("journal_all_b.txt");

        var result = PublishJournal.RecoverAll(TestDir, [TestDir], Capability());

        Assert.Contains(first.Destination, result.Restored);
        Assert.Contains(second.Destination, result.Restored);
        Assert.Equal("the previous version", File.ReadAllText(first.Destination));
        Assert.Equal("the previous version", File.ReadAllText(second.Destination));
    }

    [Fact]
    public void APublishInFlight_ShouldHaveWrittenItsJournalBeforeTheFirstMove()
    {
        // Every fixture above writes its own journal, so all of them passed with the production
        // write removed — they proved recovery works and never that anything records a publish.
        // This is the half that was missing: observed from inside the move that a crash would
        // interrupt.
        var destination = CreateTestFilePath("journal_inflight.txt");
        File.WriteAllText(destination, "previous content");

        string[] journalsDuringTheMove = [];

        using (var batch = new BoundedFileBatch(4096, "a fixture", Recovery, [TestDir]))
        {
            batch.ReportDebt = _ => { };
            batch.Stage(destination, stream => stream.Write("new content"u8));

            batch.BeforeMove = (_, _) =>
            {
                if (journalsDuringTheMove.Length == 0)
                    journalsDuringTheMove =
                        Directory.GetFiles(Recovery.Directory, "*" + PublishJournal.Extension);
            };

            batch.Publish();
        }

        Assert.NotEmpty(journalsDuringTheMove);

        // And the same file is gone once the publish commits, so a completed publish leaves
        // nothing for the next start to undo.
        foreach (var journal in journalsDuringTheMove)
            Assert.False(File.Exists(journal),
                "the journal written during the publish outlived it");
    }

    [Fact]
    public void APublishWithNothingStaged_ShouldSucceedWithoutAJournal()
    {
        // A batch that stages nothing replaces nothing, so a crash could leave nothing half-done.
        // The first version of the journal read the first staged entry unconditionally and threw
        // on exactly this case — an extraction that finds no images publishes an empty batch, and
        // two PowerPoint tests went red on it.
        using var batch = new BoundedFileBatch(4096, "a fixture", Recovery, [TestDir]);

        var written = batch.Publish();

        Assert.Empty(written);
        Assert.Empty(Directory.GetFiles(Recovery.Directory, "*" + PublishJournal.Extension));
    }

    [Fact]
    public void ASuccessfulPublish_ShouldLeaveNoJournalBehind()
    {
        // The other half of the contract: if a completed publish left its journal, the next start
        // would try to undo it.
        var destination = CreateTestFilePath("journal_clean.txt");
        File.WriteAllText(destination, "previous content");

        using (var batch = new BoundedFileBatch(4096, "a fixture", Recovery, [TestDir]))
        {
            batch.Stage(destination, stream => stream.Write("new content"u8));
            batch.Publish();
        }

        Assert.Equal("new content", File.ReadAllText(destination));
        Assert.Empty(Directory.GetFiles(Recovery.Directory, "*" + PublishJournal.Extension));
    }

    [Fact]
    public void AFailedPublishThatRolledBack_ShouldLeaveNoJournalBehind()
    {
        // The rollback already put everything back, so a journal here would make the next start
        // restore backups that no longer exist.
        var first = CreateTestFilePath("journal_rollback_a.txt");
        var second = CreateTestFilePath("journal_rollback_b.txt");
        File.WriteAllText(first, "first original");
        File.WriteAllText(second, "second original");

        using (var batch = new BoundedFileBatch(4096, "a fixture", Recovery, [TestDir]))
        {
            batch.ReportDebt = _ => { };
            batch.Stage(first, stream => stream.Write("first new"u8));
            batch.Stage(second, stream => stream.Write("second new"u8));

            // Only the forward move: the rollback moves the backup back to the same destination,
            // and a seam that fired again would be failing the recovery rather than the publish.
            var failed = false;
            batch.BeforeMove = (_, to) =>
            {
                if (to != second || failed) return;

                failed = true;
                throw new IOException("the second destination is locked");
            };

            Assert.ThrowsAny<Exception>(() => batch.Publish());
        }

        Assert.Equal("first original", File.ReadAllText(first));
        Assert.Equal("second original", File.ReadAllText(second));
        Assert.Empty(Directory.GetFiles(Recovery.Directory, "*" + PublishJournal.Extension));
    }

    [Fact]
    public void AJournalKeptAndPutBack_ShouldNotBeRecoveredTwice()
    {
        // R19-REC03. The record is genuine and its signature verifies; what it is not is still
        // outstanding. Recovery deleted it after acting, and that deletion was the only thing
        // standing between a kept copy and a second restore — over a destination that had moved
        // on in the meantime.
        var directory = Directory.CreateDirectory(Path.Combine(TestDir, "replay")).FullName;

        var destination = CreateTestFilePath("replayed_destination.txt");
        var backup = destination + ".replaced-" + new string('a', 32);
        File.WriteAllText(backup, "the previous version");
        File.WriteAllText(destination, "the half-written new version");

        var journal = Path.Combine(directory, Guid.NewGuid().ToString("N")
                                              + PublishJournal.Extension);
        var kept = PublishJournal.Signed(
            new PublishJournal.Record("a fixture",
                [new PublishJournal.Entry(destination, backup)], State: PublishJournal.Publishing),
            Capability());
        PublishJournal.Write(journal, kept);

        var first = PublishJournal.RecoverAll(directory, [TestDir], Capability());
        Assert.Contains(destination, first.Restored);
        Assert.Equal("the previous version", File.ReadAllText(destination));

        // Life goes on: the destination is written again, and the backup is recreated so the
        // replay would succeed if nothing stopped it.
        File.WriteAllText(destination, "what the caller has since written");
        File.WriteAllText(backup, "the previous version");
        PublishJournal.Write(journal, kept);

        var second = PublishJournal.RecoverAll(directory, [TestDir], Capability());

        Assert.Empty(second.Restored);
        Assert.Contains(second.Refused, r =>
            r.Contains("already been recovered", StringComparison.Ordinal));
        Assert.Equal("what the caller has since written", File.ReadAllText(destination));
    }

    [Fact]
    public void AJournalKeptForARetry_ShouldStillBeThereAtTheNextStart()
    {
        // R20-REC01. Recovery keeps a journal whenever anything was refused, so the next start can
        // try again. It then wrote the transaction id to the ledger anyway — so the next start
        // found the id, called the kept journal a replay, and deleted it. The retry it was kept
        // for never happened, and nothing said so.
        var directory = Directory.CreateDirectory(Path.Combine(TestDir, "retry")).FullName;

        // A step the current trusted roots refuse: its destination sits outside them.
        var elsewhere = Directory.CreateDirectory(Path.Combine(TestDir, "elsewhere")).FullName;
        var destination = Path.Combine(elsewhere, "not_yet_ours.txt");
        var backup = destination + ".replaced-" + new string('d', 32);
        File.WriteAllText(backup, "the previous version");
        File.WriteAllText(destination, "the half-written new version");

        var journal = Path.Combine(directory, Guid.NewGuid().ToString("N")
                                              + PublishJournal.Extension);
        WriteSigned(journal, "a fixture", [new PublishJournal.Entry(destination, backup)]);

        var narrowRoots = new[] { Path.Combine(TestDir, "retry") };

        var first = PublishJournal.RecoverAll(directory, narrowRoots, Capability());
        Assert.Empty(first.Restored);
        Assert.NotEmpty(first.Refused);
        Assert.True(File.Exists(journal), "a refused journal was not kept for a retry");

        // The next start, same refusal. The journal must still be there, refused for the reason
        // it was refused before — not tidied away as a transaction that has "already" run.
        var second = PublishJournal.RecoverAll(directory, narrowRoots, Capability());
        Assert.True(File.Exists(journal), "the kept journal was deleted as a replay");
        Assert.DoesNotContain(second.Refused, r =>
            r.Contains("already been recovered", StringComparison.Ordinal));

        // And once the roots do allow it, the retry it was kept for actually happens.
        var third = PublishJournal.RecoverAll(directory, [TestDir], Capability());
        Assert.Contains(destination, third.Restored);
        Assert.Equal("the previous version", File.ReadAllText(destination));
        Assert.False(File.Exists(journal));
    }

    [Fact]
    public void ALedgerThatCannotBeVerified_ShouldSuspendRecoveryRatherThanEmptyIt()
    {
        // R20-REC09. An unverifiable ledger was read as an empty one, which is exactly the state an
        // attacker who could write the file would want: every replay defence off, every journal
        // fresh again.
        var directory = Directory.CreateDirectory(Path.Combine(TestDir, "bad-ledger")).FullName;

        var destination = CreateTestFilePath("suspended_destination.txt");
        var backup = destination + ".replaced-" + new string('f', 32);
        File.WriteAllText(backup, "the previous version");
        File.WriteAllText(destination, "the half-written new version");

        var journal = Path.Combine(directory, Guid.NewGuid().ToString("N") + PublishJournal.Extension);
        WriteSigned(journal, "a fixture", [new PublishJournal.Entry(destination, backup)]);

        File.WriteAllText(Path.Combine(directory, PublishJournal.LedgerName),
            "{\"Recovered\":[],\"Signature\":\"not-a-signature\",\"Generation\":0}");

        var result = PublishJournal.RecoverAll(directory, [TestDir], Capability());

        Assert.Empty(result.Restored);
        Assert.Contains(result.Refused, r => r.Contains("cannot be verified", StringComparison.Ordinal));
        Assert.Equal("the half-written new version", File.ReadAllText(destination));
        Assert.True(File.Exists(journal), "recovery acted on a journal under an untrusted ledger");
    }

    [Fact]
    public void ALegacyRoot_ShouldGetNoRecoveryStateWrittenIntoIt()
    {
        // R20-REC10. The scan of an allowlisted root wrote the claim and the ledger into that
        // root — an operator's output directory acquired this server's state files.
        var state = Directory.CreateDirectory(Path.Combine(TestDir, "private-state")).FullName;
        var legacy = Directory.CreateDirectory(Path.Combine(TestDir, "operator-output")).FullName;

        var destination = Path.Combine(legacy, "legacy_destination.txt");
        var backup = destination + ".replaced-" + new string('9', 32);
        File.WriteAllText(backup, "the previous version");
        File.WriteAllText(destination, "the half-written new version");

        var journal = Path.Combine(legacy, Guid.NewGuid().ToString("N") + PublishJournal.Extension);
        var legacyStaging = Path.Combine(legacy, "legacy-journal.json");
        WriteSigned(legacyStaging, "a fixture", [new PublishJournal.Entry(destination, backup)]);
        File.Move(legacyStaging, journal);

        var result = PublishJournal.RecoverAll(legacy, [TestDir], Capability(),
            new PublishJournal.RecoveryBudget(), state);

        Assert.Contains(destination, result.Restored);
        Assert.False(File.Exists(Path.Combine(legacy, PublishJournal.LedgerName)),
            "the ledger was written into the legacy root");
        Assert.False(File.Exists(Path.Combine(legacy, PublishJournal.ClaimName)),
            "a claim was written into the legacy root");
        Assert.Empty(Directory.GetFileSystemEntries(legacy, "journals.*.pending*"));
        Assert.True(File.Exists(Path.Combine(state, PublishJournal.LedgerName)),
            "the ledger did not go to the private state directory");
    }

    [Fact]
    public void ABytesBudget_ShouldStopTheScanBeforeTheFileThatWouldExceedIt()
    {
        // R20-RES01. A count budget let every one of 512 journals be four megabytes of
        // unauthenticated JSON. Three small journals and a budget that fits two: the third is
        // refused for budget, unread.
        var directory = Directory.CreateDirectory(Path.Combine(TestDir, "bytes-budget")).FullName;

        var written = new List<string>();
        for (var i = 0; i < 3; i++)
        {
            var destination = CreateTestFilePath($"budget_{i}.txt");
            var backup = destination + ".replaced-" + new string((char)('a' + i), 32);
            File.WriteAllText(backup, "previous");
            File.WriteAllText(destination, "half-written");
            var journal = Path.Combine(directory, $"{i:D2}" + new string('0', 30) + PublishJournal.Extension);
            WriteSigned(journal, "a fixture", [new PublishJournal.Entry(destination, backup)]);
            written.Add(journal);
        }

        var perJournal = new FileInfo(written[0]).Length;
        var budget = new PublishJournal.RecoveryBudget { Bytes = perJournal * 2 + perJournal / 2 };

        var result = PublishJournal.RecoverAll(directory, [TestDir], Capability(), budget, directory);

        Assert.Equal(2, result.Restored.Count);
        Assert.Contains(result.Refused, r => r.Contains("recovery budget", StringComparison.Ordinal));
        Assert.Single(Directory.GetFiles(directory, "*" + PublishJournal.Extension));
    }

    [Fact]
    public void AJournalThatExactlyFitsTheBytesBudget_ShouldBeRecovered()
    {
        // R21-REC04. `TryCharge` succeeded, the budget hit zero, `Exhausted` said so, and the
        // journal that had just been paid for was refused.
        var directory = Directory.CreateDirectory(Path.Combine(TestDir, "exact-fit")).FullName;

        var destination = CreateTestFilePath("exact_fit.txt");
        var backup = destination + ".replaced-" + new string('e', 32);
        File.WriteAllText(backup, "previous");
        File.WriteAllText(destination, "half-written");
        var journal = Path.Combine(directory, Guid.NewGuid().ToString("N") + PublishJournal.Extension);
        WriteSigned(journal, "a fixture", [new PublishJournal.Entry(destination, backup)]);

        var exact = new FileInfo(journal).Length;

        var under = PublishJournal.RecoverAll(directory, [TestDir], Capability(),
            new PublishJournal.RecoveryBudget { Bytes = exact - 1 }, directory);
        Assert.Empty(under.Restored);
        Assert.True(File.Exists(journal));

        var fits = PublishJournal.RecoverAll(directory, [TestDir], Capability(),
            new PublishJournal.RecoveryBudget { Bytes = exact }, directory);
        Assert.Contains(destination, fits.Restored);
        Assert.False(File.Exists(journal));
    }

    [Fact]
    public void ATruncatedScan_ShouldNotEvictTheTombstoneOfAJournalItNeverReached()
    {
        // R21-REC02. `present` held only the journals inside the budget; eviction treated every
        // other id as "journal gone" and dropped its tombstone first — including the one whose
        // journal was still sitting there, beyond the count.
        var directory = Directory.CreateDirectory(Path.Combine(TestDir, "truncated")).FullName;

        // Two present journals, and a budget of one.
        var ids = new List<string>();
        for (var i = 0; i < 2; i++)
        {
            var destination = CreateTestFilePath($"truncated_{i}.txt");
            var backup = destination + ".replaced-" + new string((char)('a' + i), 32);
            File.WriteAllText(backup, "previous");
            File.WriteAllText(destination, "half-written");
            var journal = Path.Combine(directory, $"{i:D2}" + new string('0', 30) + PublishJournal.Extension);
            // Not `Describe`: that stamps this very process as the writer, and a journal whose
            // writer is still running is a publish in flight, not a crash to recover.
            var record = PublishJournal.Signed(new PublishJournal.Record("a fixture",
                [new PublishJournal.Entry(destination, backup)], State: PublishJournal.Publishing), Capability());
            PublishJournal.Write(journal, record);
            ids.Add(record.TransactionId!);
        }

        // A ledger exactly at its cap with the second journal's tombstone as its *oldest* entry:
        // one more id and eviction has to drop something, and the oldest evictable goes first.
        // Exactly at cap so the setup evicts nothing (an over-cap setup evicted, and moved the
        // kept tombstone to the newest position, where no eviction would ever reach it).
        var filler = Enumerable.Range(0, PublishJournal.MaxLedgerEntries - 1)
            .Select(n => $"filler{n:D6}").ToList();
        var recovered = new List<string> { ids[1] }.Concat(filler).ToList();
        PublishJournal.WriteLedger(directory, Capability(), recovered, 1,
            new HashSet<string>(StringComparer.Ordinal));

        var budget = new PublishJournal.RecoveryBudget { Journals = 1 };
        var result = PublishJournal.RecoverAll(directory, [TestDir], Capability(), budget, directory);

        // The first journal must really have been carried out: a scan that refused everything
        // would leave the ledger untouched and pass the assertion below for nothing.
        Assert.True(result.Restored.Count == 1,
            $"restored {result.Restored.Count}; refused: {string.Join(" | ", result.Refused)}");
        Assert.Equal("previous", File.ReadAllText(CreateTestFilePath("truncated_0.txt")));

        var ledger = PublishJournal.ReadLedger(directory, Capability());
        Assert.NotNull(ledger);
        Assert.Contains(ids[1], ledger.Recovered);
    }

    [Fact]
    public void ARootThatTakesTheLastOfTheCountBudget_ShouldStillRecoverItsJournals()
    {
        // The count is taken for a root at once. Asking the whole budget "exhausted?" inside
        // that root then refused every journal of a root that used the last of the count — a
        // start-up holding exactly its budget of journals recovered none of them.
        var directory = Directory.CreateDirectory(Path.Combine(TestDir, "last-of-count")).FullName;

        var destinations = new List<string>();
        for (var i = 0; i < 2; i++)
        {
            var destination = CreateTestFilePath($"last_of_count_{i}.txt");
            var backup = destination + ".replaced-" + new string((char)('c' + i), 32);
            File.WriteAllText(backup, "previous");
            File.WriteAllText(destination, "half-written");
            WriteSigned(Path.Combine(directory, $"{i:D2}" + new string('0', 30) + PublishJournal.Extension),
                "a fixture", [new PublishJournal.Entry(destination, backup)]);
            destinations.Add(destination);
        }

        var budget = new PublishJournal.RecoveryBudget { Journals = 2 };
        var result = PublishJournal.RecoverAll(directory, [TestDir], Capability(), budget, directory);

        Assert.Equal(2, result.Restored.Count);
        Assert.All(destinations, d => Assert.Equal("previous", File.ReadAllText(d)));
        Assert.Equal(0, budget.Journals);
    }

    [Fact]
    public void AJournalThatWasNeverRecovered_ShouldStillBeRecovered()
    {
        // The control: remembering what has been done must not stop what has not.
        var directory = Directory.CreateDirectory(Path.Combine(TestDir, "not-replay")).FullName;

        var first = CreateTestFilePath("first_destination.txt");
        var firstBackup = first + ".replaced-" + new string('b', 32);
        File.WriteAllText(firstBackup, "first previous");
        File.WriteAllText(first, "first half-written");

        WriteSigned(Path.Combine(directory, Guid.NewGuid().ToString("N")
                                            + PublishJournal.Extension),
            "a fixture", [new PublishJournal.Entry(first, firstBackup)]);

        Assert.Contains(first, PublishJournal.RecoverAll(directory, [TestDir], Capability())
            .Restored);

        var second = CreateTestFilePath("second_destination.txt");
        var secondBackup = second + ".replaced-" + new string('c', 32);
        File.WriteAllText(secondBackup, "second previous");
        File.WriteAllText(second, "second half-written");

        WriteSigned(Path.Combine(directory, Guid.NewGuid().ToString("N")
                                            + PublishJournal.Extension),
            "a fixture", [new PublishJournal.Entry(second, secondBackup)]);

        Assert.Contains(second, PublishJournal.RecoverAll(directory, [TestDir], Capability())
            .Restored);
        Assert.Equal("second previous", File.ReadAllText(second));
    }

    [SkippableFact]
    public void ARecoveryRootAnotherProcessIsRecovering_ShouldBeLeftAlone()
    {
        // R19-REC04. Two processes sharing a root both enumerated and both executed, so each
        // deleted what the other was moving. The claim is a file handle, so holding it here is
        // exactly what another process holding it looks like.
        Skip.IfNot(OperatingSystem.IsWindows(),
            "A second open of a held file is refused per-process only on Windows.");

        var directory = Directory.CreateDirectory(Path.Combine(TestDir, "claimed")).FullName;

        var destination = CreateTestFilePath("claimed_destination.txt");
        var backup = destination + ".replaced-" + new string('d', 32);
        File.WriteAllText(backup, "the previous version");
        File.WriteAllText(destination, "the half-written new version");

        WriteSigned(Path.Combine(directory, Guid.NewGuid().ToString("N")
                                            + PublishJournal.Extension),
            "a fixture", [new PublishJournal.Entry(destination, backup)]);

        using (new FileStream(Path.Combine(directory, PublishJournal.ClaimName),
                   FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None))
        {
            var result = PublishJournal.RecoverAll(directory, [TestDir], Capability());

            Assert.Empty(result.Restored);
            Assert.Contains(result.Refused, r =>
                r.Contains("another process is recovering", StringComparison.Ordinal));
            Assert.Equal("the half-written new version", File.ReadAllText(destination));
        }

        // And once the claim is free the same root recovers normally.
        Assert.Contains(destination,
            PublishJournal.RecoverAll(directory, [TestDir], Capability()).Restored);
    }

    [Fact]
    public void ManyRootsTogether_ShouldNotCostMoreThanOneStartUpsBudget()
    {
        // R19-RES01. The per-root cap bounded one directory; discovery walked the recovery
        // directory and every allowlisted root, so what a start-up spent was that cap multiplied
        // by a number the operator chooses, on directories anything writing output can fill.
        var roots = new List<string>();
        for (var r = 0; r < 3; r++)
        {
            var root = Directory.CreateDirectory(Path.Combine(TestDir, $"root{r}")).FullName;
            roots.Add(root);

            for (var i = 0; i < 40; i++)
            {
                var destination = Path.Combine(root, $"d{i}.txt");
                var backup = destination + ".replaced-" + new string('f', 32);
                File.WriteAllText(backup, "previous");
                File.WriteAllText(destination, "half-written");

                WriteSigned(Path.Combine(root, Guid.NewGuid().ToString("N")
                                               + PublishJournal.Extension),
                    "a fixture", [new PublishJournal.Entry(destination, backup)]);
            }
        }

        var budget = new PublishJournal.RecoveryBudget { Journals = 50 };

        var restored = 0;
        foreach (var root in roots)
            restored += PublishJournal.RecoverAll(root, [TestDir], Capability(), budget, root).Restored
                .Count;

        Assert.True(restored <= 50,
            $"one start-up restored {restored} records against a budget of 50");
        Assert.True(budget.Exhausted, "the budget should have been spent");
    }

    [Fact]
    public void ABudgetLargeEnoughForEverything_ShouldStillRecoverItAll()
    {
        // The control: a total budget must not turn into a cap that quietly stops recovery
        // working on an ordinary installation.
        var root = Directory.CreateDirectory(Path.Combine(TestDir, "modest")).FullName;

        for (var i = 0; i < 5; i++)
        {
            var destination = Path.Combine(root, $"m{i}.txt");
            var backup = destination + ".replaced-" + new string('f', 32);
            File.WriteAllText(backup, "previous");
            File.WriteAllText(destination, "half-written");

            WriteSigned(Path.Combine(root, Guid.NewGuid().ToString("N")
                                           + PublishJournal.Extension),
                "a fixture", [new PublishJournal.Entry(destination, backup)]);
        }

        var result = PublishJournal.RecoverAll(root, [TestDir], Capability(),
            new PublishJournal.RecoveryBudget(), root);

        Assert.Equal(5, result.Restored.Count);
    }

    [Fact]
    public void AMalformedJournal_ShouldStillSpendTheBytesBudget()
    {
        // R22-REC01. Reading it cost 64 KiB of I/O and a failed parse; a budget that only
        // recorded successful parses recorded nothing, and 512 such files could ask for
        // 2 GiB of work against a 64 MiB ceiling.
        var directory = Directory.CreateDirectory(Path.Combine(TestDir, "malformed-charge")).FullName;
        File.WriteAllBytes(Path.Combine(directory, "garbage" + PublishJournal.Extension), new byte[64 * 1024]);

        var budget = new PublishJournal.RecoveryBudget { Bytes = 1024 * 1024 };
        var result = PublishJournal.RecoverAll(directory, [TestDir], Capability(), budget, directory);

        Assert.Contains(result.Refused, r => r.Contains("could not be read", StringComparison.Ordinal));
        Assert.Equal(1024 * 1024 - 64 * 1024, budget.Bytes);
    }

    [Fact]
    public void MalformedJournalsBeyondTheBudget_ShouldNotBeRead()
    {
        // R22-REC01. Three garbage journals of 64 KiB and a budget of 160 KiB: the first two are
        // read and charged, the third is longer than what is left and must not be read at all.
        var directory = Directory.CreateDirectory(Path.Combine(TestDir, "malformed-bound")).FullName;
        for (var i = 0; i < 3; i++)
            File.WriteAllBytes(Path.Combine(directory, $"{i:D2}garbage" + PublishJournal.Extension),
                new byte[64 * 1024]);

        var budget = new PublishJournal.RecoveryBudget { Bytes = 160 * 1024 };
        var result = PublishJournal.RecoverAll(directory, [TestDir], Capability(), budget, directory);

        Assert.Equal(2, result.Refused.Count(r => r.Contains("could not be read", StringComparison.Ordinal)));
        Assert.Contains(result.Refused, r => r.Contains("remaining recovery budget", StringComparison.Ordinal));
        Assert.Equal(32 * 1024, budget.Bytes);
    }

    [Fact]
    public void AJournalOneByteOverTheRemainingBudget_ShouldNotBeReadAtAll()
    {
        // R22-REC01. Refused on its length, before any byte comes in. Garbage, so the old
        // read-then-charge code shows its hand: it read and failed to parse, and said so.
        var directory = Directory.CreateDirectory(Path.Combine(TestDir, "one-byte-over")).FullName;
        var garbage = new byte[4096];
        File.WriteAllBytes(Path.Combine(directory, "garbage" + PublishJournal.Extension), garbage);

        var budget = new PublishJournal.RecoveryBudget { Bytes = garbage.Length - 1 };
        var result = PublishJournal.RecoverAll(directory, [TestDir], Capability(), budget, directory);

        var refusal = Assert.Single(result.Refused);
        Assert.Contains("remaining recovery budget", refusal, StringComparison.Ordinal);
        Assert.DoesNotContain("could not be read", refusal, StringComparison.Ordinal);
        Assert.Equal(garbage.Length - 1, budget.Bytes);
    }

    [Fact]
    public void AnUnreadableJournalStillPresent_ShouldKeepItsTombstoneThroughAnotherRecovery()
    {
        // R22-REC04. A tombstone's journal is still in the directory but cannot be read, so its
        // identity never reaches `present`; another journal recovers in the same pass and the
        // ledger is rewritten. Absence was never proven, so nothing may be evicted — restore the
        // journal's bytes later and the tombstone is what stops its replay.
        var directory = Directory.CreateDirectory(Path.Combine(TestDir, "unreadable-present")).FullName;

        const string tombstone = "tombstone-of-the-unreadable-journal";
        var filler = Enumerable.Range(0, PublishJournal.MaxLedgerEntries - 1)
            .Select(n => $"filler{n:D6}").ToList();
        PublishJournal.WriteLedger(directory, Capability(),
            new List<string> { tombstone }.Concat(filler).ToList(), 1,
            new HashSet<string>(StringComparer.Ordinal));

        File.WriteAllBytes(Path.Combine(directory, "00unreadable" + PublishJournal.Extension), new byte[512]);

        var destination = CreateTestFilePath("unreadable_present_other.txt");
        var backup = destination + ".replaced-" + new string('e', 32);
        File.WriteAllText(backup, "previous");
        File.WriteAllText(destination, "half-written");
        WriteSigned(Path.Combine(directory, "01valid" + PublishJournal.Extension),
            "a fixture", [new PublishJournal.Entry(destination, backup)]);

        var result = PublishJournal.RecoverAll(directory, [TestDir], Capability(),
            new PublishJournal.RecoveryBudget(), directory);

        Assert.True(result.Restored.Count == 1,
            $"restored {result.Restored.Count}; refused: {string.Join(" | ", result.Refused)}");
        var ledger = PublishJournal.ReadLedger(directory, Capability());
        Assert.NotNull(ledger);
        Assert.Contains(tombstone, ledger.Recovered);
    }

    [Fact]
    public void ValidJournalsBehindABacklogOfRefusedOnes_AreReachedWithinABoundedNumberOfStarts()
    {
        // R23-REC03. Six unsigned journals that are refused and kept sort first; two valid ones
        // sort last; the budget is two per start. Without a cursor every start took the same
        // first two names and the valid ones never came up.
        var directory = Directory.CreateDirectory(Path.Combine(TestDir, "backlog")).FullName;
        foreach (var i in new[] { 4, 0, 5, 1, 3, 2 })
            File.WriteAllText(Path.Combine(directory, $"{i:D2}" + new string('a', 30) + PublishJournal.Extension),
                "not a journal");

        var destinations = new List<string>();
        foreach (var i in new[] { 11, 10 })
        {
            var destination = CreateTestFilePath($"backlog_{i}.txt");
            var backup = destination + ".replaced-" + new string((char)('a' + i - 10), 32);
            File.WriteAllText(backup, "previous");
            File.WriteAllText(destination, "half-written");
            WriteSigned(Path.Combine(directory, $"{i:D2}" + new string('b', 30) + PublishJournal.Extension),
                "a fixture", [new PublishJournal.Entry(destination, backup)]);
            destinations.Add(destination);
        }

        var restored = 0;
        var startsNeeded = 0;
        for (var start = 1; start <= 6 && restored < 2; start++)
        {
            var result = PublishJournal.RecoverAll(directory, [TestDir],
                Capability(), new PublishJournal.RecoveryBudget { Journals = 2 }, directory);
            restored += result.Restored.Count;
            startsNeeded = start;
        }

        Assert.Equal(2, restored);
        Assert.True(startsNeeded <= 4, $"took {startsNeeded} starts to reach the valid journals");
        Assert.All(destinations, d => Assert.Equal("previous", File.ReadAllText(d)));
    }

    /// <summary>Legacy inventory beyond the admission cap remains reachable across starts.</summary>
    [Fact]
    public void AJournalBeyondTheInventoryCeiling_IsReachedAcrossStarts()
    {
        // R29-REC03. The cursor used to rotate only the filesystem's first ceiling-sized slice:
        // Take(ceiling) happened before sorting. Put the only valid journal immediately beyond
        // that exact slice, without assuming anything about NTFS/ext4 enumeration order.
        var recovery = RecoveryContext.For(Path.Combine(TestDir, "beyond-ceiling-host"));
        var capability = recovery.Capability
                         ?? throw new InvalidOperationException("the fixture has no recovery key");
        var directory = recovery.Directory;
        for (var i = 0; i <= PublishJournal.JournalNameCeiling; i++)
            File.WriteAllText(Path.Combine(directory,
                $"overflow_{i:D5}{PublishJournal.Extension}"), "not a journal");

        var rawOrder = Directory.EnumerateFiles(directory, "*" + PublishJournal.Extension)
            .ToList();
        Assert.Equal(PublishJournal.JournalNameCeiling + 1, rawOrder.Count);

        var destination = CreateTestFilePath("beyond_ceiling_destination.txt");
        var backup = destination + ".replaced-" + new string('d', 32);
        File.WriteAllText(backup, "previous");
        File.WriteAllText(destination, "half-written");
        PublishJournal.Write(rawOrder[PublishJournal.JournalNameCeiling], PublishJournal.Signed(
            new PublishJournal.Record("a fixture",
                [new PublishJournal.Entry(destination, backup)],
                State: PublishJournal.Publishing), capability));

        var restored = 0;
        for (var start = 0; start < 2; start++)
            restored += PublishJournal.RecoverAll(directory, [TestDir], capability,
                new PublishJournal.RecoveryBudget
                {
                    Journals = PublishJournal.JournalNameCeiling,
                    Deadline = DateTimeOffset.UtcNow.AddMinutes(2)
                }, directory).Restored.Count;

        Assert.Equal(1, restored);
        Assert.Equal("previous", File.ReadAllText(destination));

        // Public Write/Recover is a direct single-path API, not a managed discovery producer.
        // It must remain usable even while this directory's managed inventory is at capacity.
        var directDestination = CreateTestFilePath("direct_at_capacity.txt");
        var directBackup = directDestination + ".replaced-" + new string('e', 32);
        File.WriteAllText(directBackup, "direct previous");
        File.WriteAllText(directDestination, "direct half-written");
        var directJournal = Path.Combine(directory, "direct" + PublishJournal.Extension);
        PublishJournal.Write(directJournal, PublishJournal.Signed(
            new PublishJournal.Record("a direct fixture",
                [new PublishJournal.Entry(directDestination, directBackup)],
                State: PublishJournal.Publishing), capability));

        var direct = PublishJournal.Recover(directJournal, [TestDir], capability);

        Assert.Contains(directDestination, direct.Restored);

        var protectedDestination = CreateTestFilePath("admission_at_capacity.txt");
        File.WriteAllText(protectedDestination, "the caller's original");
        using var batch = new BoundedFileBatch(1024, "an admission fixture", recovery);
        batch.Stage(protectedDestination, stream => stream.Write("replacement"u8));

        var refusal = Assert.Throws<IOException>(() => batch.Publish());

        Assert.Contains("safety limit", refusal.Message, StringComparison.Ordinal);
        Assert.Equal("the caller's original", File.ReadAllText(protectedDestination));
    }

    /// <summary>A malformed marker is refused instead of crashing startup recovery.</summary>
    [Fact]
    public void APendingIndexWithoutNames_IsRefusedInsteadOfCrashingRecovery()
    {
        var directory = Directory.CreateDirectory(Path.Combine(TestDir, "invalid-pending-index"))
            .FullName;
        PublishJournal.RecoverAll(directory, [TestDir], Capability());
        var index = Assert.Single(Directory.GetDirectories(directory, "journals.*.pending"));
        var bucket = Directory.CreateDirectory(Path.Combine(index, "0000")).FullName;
        File.WriteAllText(Path.Combine(bucket, "bad.entry"), "not-a-journal");

        var result = PublishJournal.RecoverAll(directory, [TestDir], Capability());

        Assert.Contains(result.Refused, refusal =>
            refusal.Contains("pending-journal", StringComparison.Ordinal));
    }

    /// <summary>Registration remains protected until the managed journal is durable.</summary>
    [Fact]
    public async Task AConcurrentRecovery_CannotPruneAReservationBeforeTheJournalAppears()
    {
        var directory = Directory.CreateDirectory(Path.Combine(TestDir, "index-registration-race"))
            .FullName;
        var destination = CreateTestFilePath("registration_race_destination.txt");
        var backup = destination + ".replaced-" + new string('f', 32);
        File.WriteAllText(backup, "previous");
        File.WriteAllText(destination, "half-written");
        var journal = Path.Combine(directory, Guid.NewGuid().ToString("N") + PublishJournal.Extension);

        using var recoveryStarted = new ManualResetEventSlim();
        Task<PublishJournal.RecoveryResult>? concurrentRecovery = null;
        try
        {
            PublishJournal.AfterIndexRegistrationForTest = () =>
            {
                concurrentRecovery = Task.Run(() =>
                {
                    recoveryStarted.Set();
                    return PublishJournal.RecoverAll(directory, [TestDir], Capability());
                });
                Assert.True(recoveryStarted.Wait(TimeSpan.FromSeconds(5)));
                Thread.Sleep(TimeSpan.FromMilliseconds(250));
                Assert.False(concurrentRecovery.IsCompleted,
                    "recovery acquired the index gate before the registered journal appeared");
            };

            PublishJournal.WriteManaged(journal, PublishJournal.Signed(
                new PublishJournal.Record("a fixture",
                    [new PublishJournal.Entry(destination, backup)],
                    State: PublishJournal.Publishing), Capability()));
        }
        finally
        {
            PublishJournal.AfterIndexRegistrationForTest = null;
        }

        Assert.NotNull(concurrentRecovery);
        var result = await concurrentRecovery!;
        Assert.Contains(destination, result.Restored);
    }

    /// <summary>A present but invalid cursor fails closed with an operator-visible refusal.</summary>
    [Fact]
    public void AnInvalidCursor_IsReportedToTheOperator()
    {
        var directory = Directory.CreateDirectory(Path.Combine(TestDir, "invalid-cursor"))
            .FullName;
        PublishJournal.RecoverAll(directory, [TestDir], Capability());
        var cursor = Assert.Single(Directory.GetFiles(directory, "journals.*.cursor"));
        File.WriteAllText(cursor, "not-a-journal");

        var result = PublishJournal.RecoverAll(directory, [TestDir], Capability());

        Assert.Contains(result.Refused, refusal =>
            refusal.Contains("cursor", StringComparison.Ordinal));
    }

    /// <summary>Case-sensitive roots receive distinct marker indexes and cursors.</summary>
    [SkippableFact]
    public void RootsThatDifferOnlyByCase_DoNotShareRecoveryStateOnUnix()
    {
        Skip.If(OperatingSystem.IsWindows(), "Windows path identity is case-insensitive");
        var state = Directory.CreateDirectory(Path.Combine(TestDir, "case-state")).FullName;
        var upper = Directory.CreateDirectory(Path.Combine(TestDir, "CaseRoot")).FullName;
        var lower = Directory.CreateDirectory(Path.Combine(TestDir, "caseroot")).FullName;

        PublishJournal.RecoverAll(upper, [TestDir], Capability(),
            new PublishJournal.RecoveryBudget(), state);
        PublishJournal.RecoverAll(lower, [TestDir], Capability(),
            new PublishJournal.RecoveryBudget(), state);

        Assert.Equal(2, Directory.GetDirectories(state, "journals.*.pending").Length);
        Assert.Equal(2, Directory.GetFiles(state, "journals.*.cursor").Length);
    }
}
