using System.Globalization;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Cleanup;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Tests.Infrastructure;
using Microsoft.Extensions.Logging;

namespace AsposeMcpServer.Tests.Core.Cleanup;

/// <summary>
///     R10-F01 (§22.6): the queue's problem reporting has to be wired in production, not only in
///     a test that installs the callback itself.
///     <para>
///         <c>CleanupDebtQueue.OnQueueProblem</c> existed and was assigned in exactly one place —
///         the queue's own unit tests. The service that owns the queue in production never
///         attached it, so a corrupt, full or unwritable queue stayed silent on a running server
///         while the tests reported it as handled.
///     </para>
///     <para>
///         These tests drive the real <see cref="CleanupDebtService" /> and read what reached its
///         logger. Installing a callback here would prove the same thing the old tests proved, and
///         that is the thing that was not true.
///     </para>
/// </summary>
[Collection("SerialStaticSeams")]
public class CleanupDebtServiceWiringTests : TestBase
{
    /// <summary>Reports every entry the filesystem enumerator actually yields.</summary>
    /// <param name="directory">The staging area to enumerate.</param>
    /// <param name="yielded">Called once for each entry returned by the filesystem.</param>
    /// <returns>The directory's files.</returns>
    private static IEnumerable<string> CountEntries(string directory, Action yielded)
    {
        foreach (var entry in Directory.EnumerateFiles(directory))
        {
            yielded();
            yield return entry;
        }
    }

    /// <summary>Builds a service whose queue file lives in this test's own directory.</summary>
    /// <param name="recorder">The logger to give it.</param>
    /// <returns>The service and the path of its queue file.</returns>
    private (CleanupDebtService Service, string QueueFile) AService(Recorder recorder)
    {
        return AService(recorder, TestDir);
    }

    /// <summary>Builds a service on a given temp directory.</summary>
    /// <param name="recorder">The logger to give it.</param>
    /// <param name="temporaryDirectory">The host's temp root.</param>
    /// <param name="additionalArguments">Additional command-line arguments for the host.</param>
    /// <returns>The service and the path of its queue file.</returns>
    private (CleanupDebtService Service, string QueueFile) AService(Recorder recorder,
        string temporaryDirectory, params string[] additionalArguments)
    {
        var manager = new DocumentSessionManager(
            new SessionConfig { TempDirectory = temporaryDirectory });
        var arguments = new List<string> { "--allowed-path", TestDir };
        arguments.AddRange(additionalArguments);
        var config = ServerConfig.LoadFromArgs(arguments.ToArray());

        // The queue lives in the host's recovery subdirectory, not in its temp root: a bare temp
        // root is a shared namespace, and a key sitting in one can be replaced before the server
        // first reads it (R18-SEC01).
        return (new CleanupDebtService(manager, config, recorder),
            Path.Combine(RecoveryContext.For(temporaryDirectory).Directory,
                "aspose_cleanup_debts.json"));
    }

    /// <summary>Leaves behind what a publish interrupted by a crash leaves behind.</summary>
    /// <param name="recovery">The host whose directory the record goes in and whose key signs it.</param>
    /// <param name="name">A name for the destination, so two plants do not collide.</param>
    /// <returns>The destination the record says to put back.</returns>
    private string PlantAnInterruptedPublish(RecoveryContext recovery, string name)
    {
        var destination = CreateTestFilePath(name + ".txt");
        var backup = destination + ".replaced-" + Guid.NewGuid().ToString("N");
        File.WriteAllText(backup, "the previous version of " + name);
        File.WriteAllText(destination, "the half-written new version of " + name);

        var capability = recovery.Capability;
        Assert.NotNull(capability);

        PublishJournal.WriteManaged(
            Path.Combine(recovery.Directory,
                Guid.NewGuid().ToString("N") + PublishJournal.Extension),
            PublishJournal.Signed(
                new PublishJournal.Record("a fixture",
                    [new PublishJournal.Entry(destination, backup)],
                    State: PublishJournal.Publishing),
                capability));

        return destination;
    }

    [Fact]
    public async Task AQueueFileThatCannotBeParsed_ShouldReachTheServiceLogger()
    {
        var recorder = new Recorder();
        var (service, queueFile) = AService(recorder);

        File.WriteAllText(queueFile, "{ this is not json");

        // StartAsync sweeps once, which is where the unreadable queue is met.
        await service.StartAsync(CancellationToken.None);
        try
        {
            Assert.Contains(recorder.Entries, entry =>
                entry.Level == LogLevel.Warning
                && entry.Message.Contains("could not be read", StringComparison.OrdinalIgnoreCase));
        }
        finally
        {
            await service.StopAsync(CancellationToken.None);
            service.Dispose();
        }
    }

    [Fact]
    public async Task StartingUp_ShouldRecoverAPublishInterruptedByACrash()
    {
        // The service is where an interrupted publish is first seen again: nothing ran the
        // in-process rollback, so the destination is still holding the half-written version and
        // its predecessor is sitting beside it (§23.19.2).
        var recorder = new Recorder();
        var (service, _) = AService(recorder);

        var destination = CreateTestFilePath("service_recovery.txt");
        var backup = destination + ".replaced-" + Guid.NewGuid().ToString("N");
        File.WriteAllText(backup, "the previous version");
        File.WriteAllText(destination, "the half-written new version");

        // Signed, because an unsigned record is not an instruction: recovery would leave it alone,
        // which is the point of R17-S02 and is covered by its own fixture.
        var capability = Recovery.Capability;
        Assert.NotNull(capability);

        PublishJournal.WriteManaged(
            Path.Combine(Recovery.Directory, Guid.NewGuid().ToString("N") + PublishJournal.Extension),
            PublishJournal.Signed(
                new PublishJournal.Record("a fixture",
                    [new PublishJournal.Entry(destination, backup)],
                    State: PublishJournal.Publishing),
                capability));

        await service.StartAsync(CancellationToken.None);
        try
        {
            Assert.Equal("the previous version", File.ReadAllText(destination));
            Assert.Contains(recorder.Entries, entry =>
                entry.Message.Contains("did not finish", StringComparison.OrdinalIgnoreCase));
        }
        finally
        {
            await service.StopAsync(CancellationToken.None);
            service.Dispose();
        }
    }

    [Fact]
    public async Task StartingUp_ShouldNotTreatAnAllowedOutputRootAsARecoveryJournalDirectory()
    {
        // An allowed root is caller-writable output space, not recovery authority. A file that
        // merely has the journal suffix there must never enter automatic recovery discovery,
        // even when it contains a record signed with this host's capability.
        var recorder = new Recorder();
        var (service, _) = AService(recorder);

        var destination = CreateTestFilePath("legacy_allowed_root.txt");
        var backup = destination + ".replaced-" + Guid.NewGuid().ToString("N");
        File.WriteAllText(backup, "the previous version");
        File.WriteAllText(destination, "the current version");

        var capability = Recovery.Capability;
        Assert.NotNull(capability);
        var planted = Path.Combine(TestDir,
            Guid.NewGuid().ToString("N") + PublishJournal.Extension);
        PublishJournal.Write(planted,
            PublishJournal.Signed(
                new PublishJournal.Record("an allowed-root fixture",
                    [new PublishJournal.Entry(destination, backup)],
                    State: PublishJournal.Publishing),
                capability));

        await service.StartAsync(CancellationToken.None);
        try
        {
            Assert.Equal("the current version", File.ReadAllText(destination));
            Assert.True(File.Exists(backup));
            Assert.True(File.Exists(planted));
        }
        finally
        {
            await service.StopAsync(CancellationToken.None);
            service.Dispose();
        }
    }

    [Fact]
    public async Task StartingUp_ShouldRecoverAnExplicitLegacyJournalRoot()
    {
        var recorder = new Recorder();
        var legacyRoot = Directory.CreateDirectory(Path.Combine(TestDir, "legacy-journals")).FullName;
        var destination = CreateTestFilePath("explicit_legacy_recovery.txt");
        var backup = destination + ".replaced-" + Guid.NewGuid().ToString("N");
        File.WriteAllText(backup, "the previous version");
        File.WriteAllText(destination, "the interrupted version");

        var capability = Recovery.Capability;
        Assert.NotNull(capability);
        var journal = Path.Combine(legacyRoot,
            Guid.NewGuid().ToString("N") + PublishJournal.Extension);
        PublishJournal.Write(journal,
            PublishJournal.Signed(
                new PublishJournal.Record("an explicitly configured legacy fixture",
                    [new PublishJournal.Entry(destination, backup)],
                    State: PublishJournal.Publishing),
                capability));

        var (service, _) = AService(recorder, TestDir,
            "--legacy-publish-journal-root", legacyRoot);
        await service.StartAsync(CancellationToken.None);
        try
        {
            Assert.Equal("the previous version", File.ReadAllText(destination));
            Assert.False(File.Exists(journal));
            Assert.False(File.Exists(Path.Combine(legacyRoot, PublishJournal.LedgerName)));
            Assert.False(File.Exists(Path.Combine(legacyRoot, PublishJournal.ClaimName)));
            Assert.Contains(recorder.Entries, entry =>
                entry.Message.Contains("legacy", StringComparison.OrdinalIgnoreCase)
                && entry.Message.Contains("restored", StringComparison.OrdinalIgnoreCase));
        }
        finally
        {
            await service.StopAsync(CancellationToken.None);
            service.Dispose();
        }
    }

    [Fact]
    public async Task AnOrdinarySweep_ShouldNotWarnAboutTheQueue()
    {
        // The control: the wiring must not make a healthy queue noisy, or the warning above means
        // nothing.
        var recorder = new Recorder();
        var (service, _) = AService(recorder);

        await service.StartAsync(CancellationToken.None);
        try
        {
            Assert.DoesNotContain(recorder.Entries, entry =>
                entry.Level == LogLevel.Warning
                && entry.Message.Contains("queue problem", StringComparison.OrdinalIgnoreCase));
        }
        finally
        {
            await service.StopAsync(CancellationToken.None);
            service.Dispose();
        }
    }

    [Fact]
    public async Task StartingTwice_ShouldNotLeaveASecondRegistrationBehind()
    {
        // `_queue.Record` builds a new delegate each time it is read, so the second start pushed a
        // second sink and the field only remembered that one. Stopping then removed one of the
        // two and the other stayed on the process-wide stack for the life of the process, with a
        // timer still firing that nothing held (R13-F03).
        var previous = BoundedFileBatch.DebtSinkFor(Recovery.Directory);
        var (service, _) = AService(new Recorder());

        try
        {
            await service.StartAsync(CancellationToken.None);
            await service.StartAsync(CancellationToken.None);
            await service.StopAsync(CancellationToken.None);

            Assert.Same(previous, BoundedFileBatch.DebtSinkFor(Recovery.Directory));
        }
        finally
        {
            service.Dispose();
            // Nothing to put back: the registry is keyed by root and each service releases
            // its own registration when it stops.
        }
    }

    [Fact]
    public async Task DisposingWithoutStopping_ShouldStillGiveUpTheRegistration()
    {
        // A start-up that faults, or a host disposed without being stopped. Dispose only disposed
        // the timer, so the sink stayed on a static stack, outliving the service and handing
        // debts to a queue nothing was sweeping (R13-F03).
        var previous = BoundedFileBatch.DebtSinkFor(Recovery.Directory);
        var (service, _) = AService(new Recorder());

        await service.StartAsync(CancellationToken.None);
        Assert.NotSame(previous, BoundedFileBatch.DebtSinkFor(Recovery.Directory));

        service.Dispose();

        Assert.Same(previous, BoundedFileBatch.DebtSinkFor(Recovery.Directory));
    }

    [Fact]
    public async Task StoppingTwiceThenStartingAgain_ShouldLeaveExactlyOneRegistration()
    {
        // Stop has to be repeatable — the host calls it, then disposal runs — and a service that
        // has given its registration up has to be able to take it again.
        var previous = BoundedFileBatch.DebtSinkFor(Recovery.Directory);
        var (service, _) = AService(new Recorder());

        try
        {
            await service.StartAsync(CancellationToken.None);
            await service.StopAsync(CancellationToken.None);
            await service.StopAsync(CancellationToken.None);

            Assert.Same(previous, BoundedFileBatch.DebtSinkFor(Recovery.Directory));

            await service.StartAsync(CancellationToken.None);
            Assert.NotSame(previous, BoundedFileBatch.DebtSinkFor(Recovery.Directory));

            await service.StopAsync(CancellationToken.None);
            Assert.Same(previous, BoundedFileBatch.DebtSinkFor(Recovery.Directory));
        }
        finally
        {
            service.Dispose();
            // Nothing to put back: the registry is keyed by root and each service releases
            // its own registration when it stops.
        }
    }

    [Fact]
    public async Task StoppingTheSecondServiceFirst_ShouldGiveTheSinkBackToTheFirst()
    {
        // The order the old fixture did not cover, and the one the compare-and-clear could not
        // handle: the *second* host stops while the first is still running. Measured on the
        // previous binary, that left the sink null and the running host recording nothing
        // (§23.5).
        var (firstService, _) = AService(new Recorder());
        var (secondService, _) = AService(new Recorder());

        try
        {
            await firstService.StartAsync(CancellationToken.None);
            var installedByFirst = BoundedFileBatch.DebtSinkFor(Recovery.Directory);
            Assert.NotNull(installedByFirst);

            await secondService.StartAsync(CancellationToken.None);
            Assert.NotSame(installedByFirst, BoundedFileBatch.DebtSinkFor(Recovery.Directory));

            await secondService.StopAsync(CancellationToken.None);

            Assert.Same(installedByFirst, BoundedFileBatch.DebtSinkFor(Recovery.Directory));

            await firstService.StopAsync(CancellationToken.None);
            Assert.Null(BoundedFileBatch.DebtSinkFor(Recovery.Directory));
        }
        finally
        {
            firstService.Dispose();
            secondService.Dispose();
            // Nothing to put back: the registry is keyed by root and each service releases
            // its own registration when it stops.
        }
    }

    [Fact]
    public async Task StoppingOneService_ShouldNotClearAnotherServicesSink()
    {
        // `RecordDebt` is one delegate for the process. Clearing it unconditionally on shutdown
        // took away a sink a second host had installed, leaving that host recording nothing.
        var first = new Recorder();
        var second = new Recorder();

        var (firstService, _) = AService(first);
        var (secondService, _) = AService(second);

        try
        {
            await firstService.StartAsync(CancellationToken.None);
            await secondService.StartAsync(CancellationToken.None);

            var installedBySecond = BoundedFileBatch.DebtSinkFor(Recovery.Directory);
            Assert.NotNull(installedBySecond);

            // The first service goes away; the second one's sink must survive it.
            await firstService.StopAsync(CancellationToken.None);

            Assert.Same(installedBySecond, BoundedFileBatch.DebtSinkFor(Recovery.Directory));

            // And the second registration was announced rather than silent.
            Assert.Contains(second.Entries, entry =>
                entry.Message.Contains("already registered", StringComparison.OrdinalIgnoreCase));

            await secondService.StopAsync(CancellationToken.None);
            Assert.Null(BoundedFileBatch.DebtSinkFor(Recovery.Directory));
        }
        finally
        {
            firstService.Dispose();
            secondService.Dispose();
            // Nothing to put back: the registry is keyed by root and each service releases
            // its own registration when it stops.
        }
    }

    [Fact]
    public async Task TwoHostsOnDifferentTempRoots_ShouldEachRecoverOnlyTheirOwn()
    {
        // R18-ARCH01. The journal directory and the signing key were process-wide statics that
        // every StartCore rewrote, so the second host to start moved the first's records and
        // signed with a different key. Two hosts in one process is a shape this stack is built
        // for, and the order below — start, start, stop, start again — is where the global answer
        // gave the first host the second's directory.
        var rootA = Directory.CreateDirectory(Path.Combine(TestDir, "hostA")).FullName;
        var rootB = Directory.CreateDirectory(Path.Combine(TestDir, "hostB")).FullName;

        var recoveryA = RecoveryContext.For(rootA);
        var recoveryB = RecoveryContext.For(rootB);
        Assert.NotEqual(recoveryA.Directory, recoveryB.Directory);

        var firstOfA = PlantAnInterruptedPublish(recoveryA, "hostA_first");
        var onlyOfB = PlantAnInterruptedPublish(recoveryB, "hostB_only");

        var (hostA, _) = AService(new Recorder(), rootA);
        var (hostB, _) = AService(new Recorder(), rootB);

        try
        {
            await hostA.StartAsync(CancellationToken.None);
            Assert.Equal("the previous version of hostA_first", File.ReadAllText(firstOfA));

            // B's record is in B's directory and signed with B's key, so A's start left it alone.
            Assert.Equal("the half-written new version of hostB_only", File.ReadAllText(onlyOfB));

            await hostB.StartAsync(CancellationToken.None);
            Assert.Equal("the previous version of hostB_only", File.ReadAllText(onlyOfB));

            // A restarts while B is still running. This is the step the global directory broke:
            // A looked where B had last pointed it, and its own record went unrecovered.
            var secondOfA = PlantAnInterruptedPublish(recoveryA, "hostA_second");

            await hostA.StopAsync(CancellationToken.None);
            await hostA.StartAsync(CancellationToken.None);

            Assert.Equal("the previous version of hostA_second", File.ReadAllText(secondOfA));

            await hostA.StopAsync(CancellationToken.None);
            await hostB.StopAsync(CancellationToken.None);
        }
        finally
        {
            hostA.Dispose();
            hostB.Dispose();
            // Nothing to put back: the registry is keyed by root and each service releases
            // its own registration when it stops.
        }
    }

    [Fact]
    public async Task ADebtFromOneHostsRoot_ShouldReachThatHostsQueue_NotTheNewest()
    {
        // R19-REC06. `AnnounceDebts` read a process-global "most recently registered" sink, so a
        // batch publishing into the first host's root handed its debt to the second — signed with
        // the wrong key, queued under the wrong root, and due to be swept against the wrong
        // allowlist.
        var rootA = Directory.CreateDirectory(Path.Combine(TestDir, "sinkA")).FullName;
        var rootB = Directory.CreateDirectory(Path.Combine(TestDir, "sinkB")).FullName;

        var (hostA, queueA) = AService(new Recorder(), rootA);
        var (hostB, queueB) = AService(new Recorder(), rootB);

        try
        {
            await hostA.StartAsync(CancellationToken.None);

            // B starts second, so it is the one an ambient "newest sink" would hand everything to.
            await hostB.StartAsync(CancellationToken.None);

            var destination = CreateTestFilePath("routed.txt");
            File.WriteAllText(destination, "the previous version");

            using (var batch = new BoundedFileBatch(4096, "a fixture",
                       RecoveryContext.For(rootA), [TestDir]))
            {
                batch.ReportDebt = _ => { };

                // Only the backup delete fails, which is the debt this mechanism exists for.
                batch.BeforeDelete = path =>
                {
                    if (path.Contains(".replaced-", StringComparison.Ordinal))
                        throw new IOException("the previous version is locked");
                };

                batch.Stage(destination, stream =>
                    stream.Write("the new version"u8));
                batch.Publish();
            }

            var inA = File.Exists(queueA) ? File.ReadAllText(queueA) : string.Empty;
            var inB = File.Exists(queueB) ? File.ReadAllText(queueB) : string.Empty;

            Assert.Contains(".replaced-", inA, StringComparison.Ordinal);
            Assert.DoesNotContain(".replaced-", inB, StringComparison.Ordinal);

            await hostA.StopAsync(CancellationToken.None);
            await hostB.StopAsync(CancellationToken.None);
        }
        finally
        {
            hostA.Dispose();
            hostB.Dispose();
            // Nothing to put back: the registry is keyed by root and each service releases
            // its own registration when it stops.
        }
    }

    [Fact]
    public async Task StartingUp_ShouldSweepStaleInputCopiesAndLeaveFreshOnes()
    {
        // R20-CNV03. The copy's doc promised a start-up sweep of `authorised-inputs` before any
        // sweep existed.
        var staging = Directory.CreateDirectory(
            Path.Combine(Recovery.Directory, ImmutableInputCopy.DirectoryName)).FullName;

        var stale = Path.Combine(staging, "stale." + Guid.NewGuid().ToString("N") + ".mht");
        var fresh = Path.Combine(staging, "fresh." + Guid.NewGuid().ToString("N") + ".mht");
        File.WriteAllText(stale, "left by a crash");
        File.WriteAllText(fresh, "in use by another host right now");
        File.SetLastWriteTimeUtc(stale, DateTime.UtcNow.AddHours(-2));

        var (service, _) = AService(new Recorder());
        await service.StartAsync(CancellationToken.None);
        try
        {
            Assert.False(File.Exists(stale), "a stale input copy survived the start-up sweep");
            Assert.True(File.Exists(fresh), "a fresh input copy was swept out from under a conversion");
        }
        finally
        {
            await service.StopAsync(CancellationToken.None);
            service.Dispose();
        }
    }

    [Fact]
    public void TwoHostsOnOneRootWithDifferentAllowlists_ShouldEachGetTheirOwnDebts()
    {
        // R20-REC06. Keyed by root alone, these two were the same host to the registry, and a
        // batch running under the first allowlist could hand its debt to the second.
        var previous = BoundedFileBatch.RecordDebt;
        var first = Directory.CreateDirectory(Path.Combine(TestDir, "allow-first")).FullName;
        var second = Directory.CreateDirectory(Path.Combine(TestDir, "allow-second")).FullName;

        var firstDebts = new List<string>();
        var secondDebts = new List<string>();

        // The same instances are removed in the finally: the registry compares delegates by
        // reference, and an earlier version built new lambdas there, so nothing was ever removed
        // and the sinks leaked into every later test (R21-TST01).
        Action<string, string> firstSink = (path, _) => firstDebts.Add(path);
        Action<string, string> secondSink = (path, _) => secondDebts.Add(path);
        BoundedFileBatch.AddDebtSink(Recovery.Directory, [first], firstSink);
        BoundedFileBatch.AddDebtSink(Recovery.Directory, [second], secondSink);

        try
        {
            var forFirst = BoundedFileBatch.DebtSinkFor(Recovery.Directory, [first]);
            var forSecond = BoundedFileBatch.DebtSinkFor(Recovery.Directory, [second]);
            Assert.NotNull(forFirst);
            Assert.NotNull(forSecond);
            Assert.NotSame(forFirst, forSecond);

            forFirst("a", "locked");
            Assert.Equal(["a"], firstDebts);
            Assert.Empty(secondDebts);

            // No host deletes under a third allowlist; nobody is "close enough".
            var third = Directory.CreateDirectory(Path.Combine(TestDir, "allow-third")).FullName;
            Assert.Null(BoundedFileBatch.DebtSinkFor(Recovery.Directory, [third]));
        }
        finally
        {
            BoundedFileBatch.RemoveDebtSink(Recovery.Directory, firstSink);
            BoundedFileBatch.RemoveDebtSink(Recovery.Directory, secondSink);
            BoundedFileBatch.RecordDebt = previous;

            Assert.Null(BoundedFileBatch.DebtSinkFor(Recovery.Directory, [first]));
            Assert.Null(BoundedFileBatch.DebtSinkFor(Recovery.Directory, [second]));
        }
    }

    [Fact]
    public async Task StartingUp_ShouldReachStaleCopiesBehindManyFreshOnes_AndLeaveOneInUse()
    {
        // R21-REC05. Three hundred fresh copies ahead of one stale one starved it for ever; and a
        // stale copy somebody still holds open is not a crash leftover. Named so the stale file
        // enumerates *after* the fresh ones: a first-256-names sweep must genuinely never reach it.
        var staging = Directory.CreateDirectory(
            Path.Combine(Recovery.Directory, ImmutableInputCopy.DirectoryName)).FullName;

        var stale = Path.Combine(staging, "zzz-stale." + Guid.NewGuid().ToString("N") + ".mht");
        File.WriteAllText(stale, "left by a crash");
        File.SetLastWriteTimeUtc(stale, DateTime.UtcNow.AddHours(-3));

        for (var i = 0; i < 300; i++)
            File.WriteAllText(Path.Combine(staging, $"aaa-fresh-{i:D3}.{Guid.NewGuid():N}.mht"), "in use");

        var held = Path.Combine(staging, "held." + Guid.NewGuid().ToString("N") + ".mht");
        File.WriteAllText(held, "still being converted");
        File.SetLastWriteTimeUtc(held, DateTime.UtcNow.AddHours(-3));

        var (service, _) = AService(new Recorder());
        await using (new FileStream(held, FileMode.Open, FileAccess.Read, FileShare.None))
        {
            await service.StartAsync(CancellationToken.None);
        }

        try
        {
            Assert.False(File.Exists(stale), "the stale copy behind the fresh ones was never reached");
            Assert.True(File.Exists(held), "a copy still held open was swept");
        }
        finally
        {
            await service.StopAsync(CancellationToken.None);
            service.Dispose();
        }
    }

    [Fact]
    public async Task StartingUp_ShouldLeaveAStaleCopyWhoseConversionIsStillRunning_EvenWithNoHandleOnIt()
    {
        // R22-REC02. Production holds no handle on its copy between the scan and the loader; an
        // instant's exclusive open found nothing there and a sweep in another host deleted a
        // running conversion's input. The lease is held for the copy's whole lifetime, and that
        // is what the sweep must fail to take.
        var source = CreateTestFilePath("running_conversion_input.mht");
        File.WriteAllText(source, "the caller's input");

        var copy = ImmutableInputCopy.Of(source, Recovery, [TestDir]);
        File.SetLastWriteTimeUtc(copy.Path, DateTime.UtcNow.AddHours(-3));
        Assert.True(File.Exists(ImmutableInputCopy.LeasePathOf(copy.Path)), "a copy is created leased");

        var (service, _) = AService(new Recorder());
        try
        {
            await service.StartAsync(CancellationToken.None);
            Assert.True(File.Exists(copy.Path), "a leased copy with no handle on it was swept");
        }
        finally
        {
            await service.StopAsync(CancellationToken.None);
            service.Dispose();
        }

        copy.Dispose();
        Assert.False(File.Exists(copy.Path), "disposal did not remove the copy");
        Assert.False(File.Exists(ImmutableInputCopy.LeasePathOf(copy.Path)), "disposal left the lease behind");
    }

    [Fact]
    public async Task StartingUp_ShouldRemoveAStaleCopyWhoseLeaseDiedWithItsProcess()
    {
        // The crash shape: the lease closed with the process and deleted itself, the copy stayed.
        var staging = Directory.CreateDirectory(
            Path.Combine(Recovery.Directory, ImmutableInputCopy.DirectoryName)).FullName;
        var abandoned = Path.Combine(staging, "abandoned." + Guid.NewGuid().ToString("N") + ".mht");
        File.WriteAllText(abandoned, "left by a crash");
        File.SetLastWriteTimeUtc(abandoned, DateTime.UtcNow.AddHours(-3));

        var (service, _) = AService(new Recorder());
        try
        {
            await service.StartAsync(CancellationToken.None);
            Assert.False(File.Exists(abandoned), "an unleased stale copy was not swept");
            Assert.False(File.Exists(ImmutableInputCopy.LeasePathOf(abandoned)), "the sweep left its own lease behind");
        }
        finally
        {
            await service.StopAsync(CancellationToken.None);
            service.Dispose();
        }
    }

    [Fact]
    public async Task StartingUp_ShouldRemoveAStaleCopyFromANonceShard()
    {
        var staging = Directory.CreateDirectory(
            Path.Combine(Recovery.Directory, ImmutableInputCopy.DirectoryName)).FullName;
        var shard = Directory.CreateDirectory(Path.Combine(staging, "7f")).FullName;
        var abandoned = Path.Combine(shard, "abandoned." + Guid.NewGuid().ToString("N") + ".mht");
        File.WriteAllText(abandoned, "left by a crash");
        File.SetLastWriteTimeUtc(abandoned, DateTime.UtcNow.AddHours(-3));

        var (service, _) = AService(new Recorder());
        try
        {
            await service.StartAsync(CancellationToken.None);
            Assert.False(File.Exists(abandoned), "a stale copy inside a nonce shard was not swept");
        }
        finally
        {
            await service.StopAsync(CancellationToken.None);
            service.Dispose();
        }
    }

    [Fact]
    public async Task StartingUp_ShouldNotLetANonDeletableBucketHideTheNextBucketInTheSameShard()
    {
        // R29-REC02 review correction. A cursor for only the first nonce byte could replay a
        // bucket's non-deletable prefix forever. Exact second-byte buckets let it advance without
        // asking that prefix's enumerator to yield it again.
        var staging = Directory.CreateDirectory(
            Path.Combine(Recovery.Directory, ImmutableInputCopy.DirectoryName)).FullName;
        var first = Directory.CreateDirectory(Path.Combine(staging, "7f", "00")).FullName;
        var second = Directory.CreateDirectory(Path.Combine(staging, "7f", "01")).FullName;
        File.WriteAllText(Path.Combine(first, "input.mht"), "still young");
        var stale = Path.Combine(second, "input.mht");
        File.WriteAllText(stale, "left by a crash");
        File.SetLastWriteTimeUtc(stale, DateTime.UtcNow.AddHours(-3));

        // Area zero is flat, 1..256 are legacy shards, then 256 buckets per first byte.
        File.WriteAllText(Path.Combine(Recovery.Directory, "staged-inputs.cursor"),
            (257 + 0x7f * 256).ToString(CultureInfo.InvariantCulture));
        var ceiling = CleanupDebtService.EnumerationCeiling;
        CleanupDebtService.EnumerationCeiling = 4;
        try
        {
            var (service, _) = AService(new Recorder());
            await service.StartAsync(CancellationToken.None);
            await service.StopAsync(CancellationToken.None);
            service.Dispose();
        }
        finally
        {
            CleanupDebtService.EnumerationCeiling = ceiling;
        }

        Assert.False(File.Exists(stale), "the next bucket was hidden by a young copy");
    }

    [Fact]
    public async Task StartingUp_ShouldExamineAtMostTheCeiling_AndStillReachEveryStaleCopyAcrossStarts()
    {
        // R23-REC02. A hundred stale copies, a ceiling of forty: each start examines at most
        // forty entries, and three starts remove them all, because each continues where the
        // last stopped.
        var staging = Directory.CreateDirectory(
            Path.Combine(Recovery.Directory, ImmutableInputCopy.DirectoryName)).FullName;
        for (var i = 0; i < 100; i++)
        {
            var stale = Path.Combine(staging, $"stale-{i:D3}.{Guid.NewGuid():N}.mht");
            File.WriteAllText(stale, "left by a crash");
            File.SetLastWriteTimeUtc(stale, DateTime.UtcNow.AddHours(-3));
        }

        var examinedPerSweep = new List<int>();
        var ceiling = CleanupDebtService.EnumerationCeiling;
        CleanupDebtService.EnumerationCeiling = 40;
        CleanupDebtService.Examined = n => examinedPerSweep.Add(n);
        try
        {
            for (var sweep = 0; sweep < 3; sweep++)
            {
                var (service, _) = AService(new Recorder());
                await service.StartAsync(CancellationToken.None);
                await service.StopAsync(CancellationToken.None);
                service.Dispose();
            }
        }
        finally
        {
            CleanupDebtService.EnumerationCeiling = ceiling;
            CleanupDebtService.Examined = null;
        }

        Assert.Equal(3, examinedPerSweep.Count);
        Assert.All(examinedPerSweep, n => Assert.True(n <= 40, $"a sweep examined {n} entries"));
        Assert.Empty(Directory.GetFiles(staging, "*.mht"));
    }

    [Fact]
    public async Task StartingUp_ShouldChargeEveryEnumeratorYieldToTheCeiling()
    {
        // R29-REC02. The old ordinal cursor reported only entries after its skip as examined;
        // reaching that cursor replayed 40, then 80, then 100 filesystem yields under a stated
        // ceiling of 40. This counts the real iterator rather than the service's own metric.
        var staging = Directory.CreateDirectory(
            Path.Combine(Recovery.Directory, ImmutableInputCopy.DirectoryName)).FullName;
        for (var i = 0; i < 100; i++)
            File.WriteAllText(Path.Combine(staging, $"young-{i:D3}.mht"), "still active");

        var yielded = 0;
        var actualPerSweep = new List<int>();
        var ceiling = CleanupDebtService.EnumerationCeiling;
        var enumerate = CleanupDebtService.EnumerateStagedInputFiles;
        CleanupDebtService.EnumerationCeiling = 40;
        CleanupDebtService.EnumerateStagedInputFiles = directory =>
            CountEntries(directory, () => yielded++);
        CleanupDebtService.Examined = _ =>
        {
            actualPerSweep.Add(yielded);
            yielded = 0;
        };

        try
        {
            for (var sweep = 0; sweep < 3; sweep++)
            {
                var (service, _) = AService(new Recorder());
                await service.StartAsync(CancellationToken.None);
                await service.StopAsync(CancellationToken.None);
                service.Dispose();
            }
        }
        finally
        {
            CleanupDebtService.EnumerationCeiling = ceiling;
            CleanupDebtService.EnumerateStagedInputFiles = enumerate;
            CleanupDebtService.Examined = null;
        }

        Assert.Equal(3, actualPerSweep.Count);
        Assert.All(actualPerSweep, actual =>
            Assert.InRange(actual, 1, 40));
    }

    /// <summary>Collects everything logged, so a fixture can assert on it.</summary>
    private sealed class Recorder : ILogger<CleanupDebtService>
    {
        /// <summary>Every message logged, with its level.</summary>
        public List<(LogLevel Level, string Message)> Entries { get; } = [];

        /// <inheritdoc />
        public IDisposable? BeginScope<TState>(TState state) where TState : notnull
        {
            return null;
        }

        /// <inheritdoc />
        public bool IsEnabled(LogLevel logLevel)
        {
            return true;
        }

        /// <inheritdoc />
        public void Log<TState>(LogLevel logLevel, EventId eventId, TState state, Exception? exception,
            Func<TState, Exception?, string> formatter)
        {
            Entries.Add((logLevel, formatter(state, exception)));
        }
    }
}
