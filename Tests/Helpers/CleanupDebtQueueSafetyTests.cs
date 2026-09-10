using System.Diagnostics;
using System.Text.Json;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Helpers;

/// <summary>
///     R9-F01 (§21.2.2): the queue is a record of intent, never an authority to delete.
///     <para>
///         Its containment check looked only at the leaf: a directory on the path could be
///         replaced with a junction after the debt was recorded and the delete would follow it out
///         of the allowlist, because the leaf itself is not a link. The allowlist comparison was
///         also fixed to <c>OrdinalIgnoreCase</c>, which accepts a path a case-sensitive
///         filesystem would place outside the root.
///     </para>
///     <para>
///         And a queue file that cannot be parsed was treated as an empty one, so the next sweep
///         wrote over it — losing every debt recorded before the corruption rather than reporting
///         that the record is unreadable.
///     </para>
/// </summary>
public class CleanupDebtQueueSafetyTests : TestBase
{
    /// <summary>A queue rooted at the test directory.</summary>
    /// <param name="name">Queue file name.</param>
    /// <param name="roots">Allowlisted roots; the test directory when omitted.</param>
    /// <returns>The queue.</returns>
    private CleanupDebtQueue AQueue(string name, IReadOnlyList<string>? roots = null)
    {
        return new CleanupDebtQueue(CreateTestFilePath(name), roots ?? [TestDir], Recovery,
            CleanupDebtQueue.DefaultRetention);
    }

    /// <summary>Creates a directory junction, which needs no elevation on Windows.</summary>
    /// <param name="link">Where the junction goes.</param>
    /// <param name="target">What it points at.</param>
    /// <returns><c>true</c> when the junction was created.</returns>
    private static bool TryCreateJunction(string link, string target)
    {
        if (!OperatingSystem.IsWindows()) return false;

        var process = Process.Start(new ProcessStartInfo("cmd.exe",
            $"/c mklink /J \"{link}\" \"{target}\"")
        {
            CreateNoWindow = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true
        });

        process?.WaitForExit(10_000);
        return process?.ExitCode == 0 && Directory.Exists(link);
    }

    [SkippableFact]
    public void AnAncestorReplacedByAJunction_ShouldStopTheDelete()
    {
        Skip.IfNot(OperatingSystem.IsWindows(), "Directory junctions are a Windows construct.");

        // The shape: the debt is recorded under an allowlisted directory, and that directory is
        // then replaced by a junction pointing outside it. The leaf is not a link, so a check
        // that looks only at the leaf sees nothing wrong.
        var inside = Path.Combine(TestDir, "reaper_inside");
        var outside = Path.Combine(TestDir, "reaper_outside");
        Directory.CreateDirectory(inside);
        Directory.CreateDirectory(outside);

        var recorded = Path.Combine(inside, "debt.txt");
        File.WriteAllText(recorded, "queued");

        var queue = AQueue("junction_queue.json", [inside]);
        queue.Record(recorded, "locked");

        var decoy = Path.Combine(outside, "debt.txt");
        File.WriteAllText(decoy, "must survive");

        Directory.Delete(inside, true);
        Skip.IfNot(TryCreateJunction(inside, outside),
            "This account cannot create a directory junction.");

        var swept = queue.Sweep();

        Assert.True(File.Exists(decoy),
            "the sweep followed a junction and deleted a file outside the allowlisted root");
        Assert.Contains(recorded, swept.Refused);
    }

    [SkippableFact]
    public void AnAllowlistedRootReachedThroughAJunction_ShouldStillBeRefused()
    {
        Skip.IfNot(OperatingSystem.IsWindows(), "Directory junctions are a Windows construct.");

        // The same hazard from the other side: the recorded path is spelled through a junction
        // that happens to land inside the root. The queue approved a real path, not this one.
        var real = Path.Combine(TestDir, "reaper_real");
        Directory.CreateDirectory(real);

        var target = Path.Combine(real, "through_link.txt");
        File.WriteAllText(target, "queued");

        var link = Path.Combine(TestDir, "reaper_link");
        Skip.IfNot(TryCreateJunction(link, real),
            "This account cannot create a directory junction.");

        var queue = AQueue("through_link_queue.json", [TestDir]);
        queue.Record(Path.Combine(link, "through_link.txt"), "locked");

        var swept = queue.Sweep();

        Assert.Empty(swept.Deleted);
        Assert.True(File.Exists(target),
            "the sweep deleted through a junction rather than refusing the spelling it never approved");
    }

    [Fact]
    public void AQueueFileThatCannotBeParsed_ShouldNotBeSilentlyReplaced()
    {
        // Treating a corrupt queue as an empty one means the next write destroys every debt
        // recorded before the corruption — the queue quietly forgets what it was for.
        var queueFile = CreateTestFilePath("corrupt_queue.json");
        var queue = new CleanupDebtQueue(queueFile, [TestDir], Recovery, CleanupDebtQueue.DefaultRetention);

        var target = CreateTestFilePath("corrupt_target.txt");
        File.WriteAllText(target, "queued");
        queue.Record(target, "locked");

        Assert.NotEmpty(queue.Pending());

        File.WriteAllText(queueFile, "{ this is not json");

        var problems = new List<string>();
        queue.OnQueueProblem = problems.Add;

        queue.Sweep();

        Assert.NotEmpty(problems);
        Assert.Contains(problems, message =>
            message.Contains("could not be read", StringComparison.OrdinalIgnoreCase));

        // And the unreadable file is preserved rather than overwritten with an empty list.
        Assert.Contains("this is not json", File.ReadAllText(queueFile), StringComparison.Ordinal);
    }

    [Fact]
    public void AQueueThatIsFull_ShouldSaySoRatherThanDropTheDebt()
    {
        var queue = AQueue("full_queue.json");
        var problems = new List<string>();
        queue.OnQueueProblem = problems.Add;

        for (var i = 0; i < CleanupDebtQueue.MaxDebts + 5; i++)
        {
            var path = CreateTestFilePath($"full_{i}.txt");
            File.WriteAllText(path, "x");
            queue.Record(path, "locked");
        }

        Assert.Equal(CleanupDebtQueue.MaxDebts, queue.Pending().Count);
        Assert.NotEmpty(problems);
        Assert.Contains(problems, message =>
            message.Contains("full", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void AQueueThatCannotBeWritten_ShouldSaySo()
    {
        // A debt that is never persisted is a debt nobody will ever act on, and the caller was
        // told nothing about it.
        var directory = Path.Combine(TestDir, "unwritable_queue");
        Directory.CreateDirectory(directory);

        var queueFile = Path.Combine(directory, "queue.json");
        var queue = new CleanupDebtQueue(queueFile, [TestDir], Recovery, CleanupDebtQueue.DefaultRetention);

        var problems = new List<string>();
        queue.OnQueueProblem = problems.Add;
        queue.WriteText = (_, _) => throw new IOException("the queue file is locked");

        var target = CreateTestFilePath("unwritable_target.txt");
        File.WriteAllText(target, "queued");
        queue.Record(target, "locked");

        Assert.NotEmpty(problems);
        Assert.Contains(problems, message =>
            message.Contains("could not be written", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void AnOrdinaryDebt_ShouldStillBeSweptWithoutComplaint()
    {
        // The control: none of the guards above may make the ordinary case noisy or refuse it.
        var queue = AQueue("ordinary_queue.json");
        var problems = new List<string>();
        queue.OnQueueProblem = problems.Add;

        var target = CreateTestFilePath("ordinary_target.txt");
        File.WriteAllText(target, "queued");
        queue.Record(target, "locked");

        var swept = queue.Sweep();

        Assert.Contains(target, swept.Deleted);
        Assert.Empty(swept.Refused);
        Assert.Empty(problems);
        Assert.False(File.Exists(target));
    }

    [Fact]
    public void TheQueueFileItself_ShouldRoundTripThroughJson()
    {
        // Guards that refuse everything would pass most of the tests above; this pins that a
        // recorded debt is still readable back.
        var queueFile = CreateTestFilePath("roundtrip_queue.json");
        var queue = new CleanupDebtQueue(queueFile, [TestDir], Recovery, CleanupDebtQueue.DefaultRetention);

        var target = CreateTestFilePath("roundtrip_target.txt");
        File.WriteAllText(target, "queued");
        queue.Record(target, "locked");

        var raw = JsonSerializer.Deserialize<List<JsonElement>>(File.ReadAllText(queueFile));

        Assert.NotNull(raw);
        Assert.Single(raw);
    }

    [Fact]
    public void ADebtPlantedInTheQueueFile_ShouldNotBeDeleted()
    {
        // R17-S01. The queue is a JSON file in the temp directory, and a `Path` read out of it was
        // a path to delete: canonical, inside the allowlist if there is one, not a link. With no
        // allowlist — the default — even that check is skipped. Containment is not provenance, and
        // an entry naming an unrelated file is indistinguishable from one this queue wrote.
        var victim = CreateTestFilePath("someone_elses_output.docx");
        File.WriteAllText(victim, "a document this server produced earlier");

        var queueFile = CreateTestFilePath("planted_queue.json");
        var now = DateTimeOffset.UtcNow;
        File.WriteAllText(queueFile, JsonSerializer.Serialize(new[]
        {
            new
            {
                Path = victim, Attempts = 0, FirstSeenUtc = now, NextAttemptUtc = now,
                LastError = "planted"
            }
        }));

        var queue = new CleanupDebtQueue(queueFile, [], Recovery);
        var result = queue.Sweep();

        Assert.DoesNotContain(victim, result.Deleted);
        Assert.True(File.Exists(victim), "a planted queue entry deleted an unrelated file");
        Assert.Equal("a document this server produced earlier", File.ReadAllText(victim));
    }

    [Fact]
    public void ADebtThisQueueRecorded_ShouldStillBeSwept()
    {
        // The control: signing must not stop the queue doing its job.
        var target = CreateTestFilePath("recorded_and_swept.txt");
        File.WriteAllText(target, "content");

        var queue = new CleanupDebtQueue(CreateTestFilePath("recorded_queue.json"), [TestDir], Recovery);
        queue.Record(target, "locked");

        Assert.Contains(target, queue.Sweep().Deleted);
        Assert.False(File.Exists(target));
    }

    [Fact]
    public void AQueueFileLargerThanAQueueMayBe_ShouldBeRefusedWithoutBeingParsed()
    {
        // R17-S03. `MaxDebts` bounded what this process records and nothing bounded what it reads
        // back, so the work one start-up did was chosen by whoever could write the file. The
        // journal has had this cap since R13-SEC01; the file next door never got one.
        var queueFile = CreateTestFilePath("huge_queue.json");
        File.WriteAllText(queueFile, new string('x', 9 * 1024 * 1024));

        var problems = new List<string>();
        var queue = new CleanupDebtQueue(queueFile, [TestDir], Recovery) { OnQueueProblem = problems.Add };

        var result = queue.Sweep();

        Assert.Empty(result.Deleted);
        Assert.Contains(problems, problem =>
            problem.Contains("larger than a queue may be", StringComparison.Ordinal));
    }

    [Fact]
    public void AQueueHoldingMoreEntriesThanItRecords_ShouldBeRefused()
    {
        var queueFile = CreateTestFilePath("overfull_queue.json");
        var now = DateTimeOffset.UtcNow;
        var entries = Enumerable.Range(0, CleanupDebtQueue.MaxDebts + 1).Select(i => new
        {
            Path = CreateTestFilePath($"overfull_{i}.txt"),
            Attempts = 0, FirstSeenUtc = now, NextAttemptUtc = now, LastError = "planted"
        });
        File.WriteAllText(queueFile, JsonSerializer.Serialize(entries));

        var problems = new List<string>();
        var queue = new CleanupDebtQueue(queueFile, [TestDir], Recovery) { OnQueueProblem = problems.Add };

        Assert.Empty(queue.Sweep().Deleted);
        Assert.Contains(problems, problem =>
            problem.Contains("more than the", StringComparison.Ordinal));
    }

    [Fact]
    public void AUncDebt_ShouldBeRefusedBeforeTheFilesystemIsAsked()
    {
        // R17-S04. `MayDelete` handed the path straight to FileInfo/DirectoryInfo, and on Windows
        // a lookup on an attacker-named UNC path is an SMB request that the refusal afterwards
        // cannot take back. The journal has refused these shapes before touching the filesystem
        // since R13-SEC01; the queue never did.
        var asked = new List<string>();
        var queueFile = CreateTestFilePath("unc_queue.json");
        var now = DateTimeOffset.UtcNow;
        File.WriteAllText(queueFile, JsonSerializer.Serialize(new[]
        {
            new
            {
                Path = @"\\attacker.invalid\share\payload.txt", Attempts = 0,
                FirstSeenUtc = now, NextAttemptUtc = now, LastError = "planted"
            }
        }));

        var queue = new CleanupDebtQueue(queueFile, [], Recovery)
        {
            DeleteIf = (p, _) =>
            {
                asked.Add(p);
                return true;
            }
        };

        var result = queue.Sweep();

        Assert.Empty(result.Deleted);
        Assert.Empty(asked);
    }

    [Fact]
    public void AKeptDebtReplayedAfterThePathIsReused_ShouldNotDeleteTheNewFile()
    {
        // R18-SEC04. The signature says this server recorded a debt about that path once. It says
        // nothing about what is at the path now, so a legitimate record, kept and put back after
        // the name was reused, deleted whoever had taken it.
        var target = CreateTestFilePath("reused_path.txt");
        File.WriteAllText(target, "the file the debt was about");

        var queueFile = CreateTestFilePath("replay_queue.json");
        AQueue("replay_queue.json").Record(target, "locked");

        // The record as it stands: legitimately signed, about a file that is about to go.
        var kept = File.ReadAllText(queueFile);

        File.Delete(target);
        File.WriteAllText(target, "somebody else's file, at the same name");
        File.WriteAllText(queueFile, kept);

        var result = AQueue("replay_queue.json").Sweep();

        Assert.DoesNotContain(target, result.Deleted);
        Assert.Contains(target, result.Refused);
        Assert.Equal("somebody else's file, at the same name", File.ReadAllText(target));
    }

    [Fact]
    public void ADebtAboutTheFileItWasRecordedAgainst_ShouldStillBeSwept()
    {
        // The control: binding a debt to its target must not stop the queue doing its job.
        var target = CreateTestFilePath("unchanged_target.txt");
        File.WriteAllText(target, "content");

        var queue = AQueue("unchanged_queue.json");
        queue.Record(target, "locked");

        Assert.Contains(target, queue.Sweep().Deleted);
        Assert.False(File.Exists(target));
    }

    [Fact]
    public void TheStagingFileAQueueWriteUses_ShouldNotBeAPredictableName()
    {
        // R18-SEC02. `<queue>.writing` is a name anyone watching can predict, and a path-based
        // write to a name someone has left a link at follows the link.
        var queueFile = CreateTestFilePath("staging_name_queue.json");
        var seen = new List<string>();

        var queue = new CleanupDebtQueue(queueFile, [TestDir], Recovery)
        {
            WriteText = (stream, content) =>
            {
                seen.Add(((FileStream)stream).Name);
                using var writer = new StreamWriter(stream, leaveOpen: true);
                writer.Write(content);
            }
        };

        var target = CreateTestFilePath("staging_name_target.txt");
        File.WriteAllText(target, "content");
        queue.Record(target, "locked");

        Assert.NotEmpty(seen);
        Assert.All(seen, path =>
        {
            Assert.NotEqual(queueFile + ".writing", path);
            Assert.StartsWith(queueFile + ".writing-", path, StringComparison.Ordinal);
        });
    }

    [Fact]
    public void ADebtOutsideOneHostsAllowlist_ShouldSurviveThatHostsSweepForTheHostItBelongsTo()
    {
        // R20-REC06. Two hosts on one recovery root share one queue file. The one whose allowlist
        // did not cover a debt refused it — and then wrote the remaining list back without it, so
        // the host whose allowlist *did* cover it never saw it again.
        var ours = Directory.CreateDirectory(Path.Combine(TestDir, "ours")).FullName;
        var theirs = Directory.CreateDirectory(Path.Combine(TestDir, "theirs")).FullName;

        var target = Path.Combine(ours, "left_behind.txt");
        File.WriteAllText(target, "content");

        var queueFile = CreateTestFilePath("shared_queue.json");
        var narrow = new CleanupDebtQueue(queueFile, [theirs], Recovery);
        var covering = new CleanupDebtQueue(queueFile, [ours], Recovery);

        covering.Record(target, "locked");

        var narrowSweep = narrow.Sweep();
        Assert.Contains(target, narrowSweep.Refused);
        Assert.True(File.Exists(target), "the narrow host deleted a file outside its allowlist");

        // Still there for the host it belongs to.
        Assert.Contains(target, covering.Pending().Select(d => d.Path));
        Assert.Contains(target, covering.Sweep().Deleted);
        Assert.False(File.Exists(target));
    }

    [Fact]
    public void AnErrorLongerThanADebtMayCarry_ShouldBeBoundedRatherThanLoseTheDebt()
    {
        // R20-REC11. The error was signed in at any length; the read side then refused the whole
        // debt as malformed, so the cleanup it tracked was silently forgotten.
        var target = CreateTestFilePath("long_error_target.txt");
        File.WriteAllText(target, "content");

        var queue = AQueue("long_error_queue.json");
        queue.Record(target, new string('e', 10_000));

        var pending = Assert.Single(queue.Pending());
        Assert.Equal(target, pending.Path);
        Assert.True(pending.LastError.Length <= 2048);
    }

    [Fact]
    public void ARetryWhoseErrorIsTooLong_ShouldStillBeADebtAtTheNextSweep()
    {
        // R21-REC01. `Record` bounded the error; the retry path signed `ex.Message` as it came.
        // The next read refused the whole debt as malformed, and the cleanup was forgotten.
        var target = CreateTestFilePath("retry_long_error.txt");
        File.WriteAllText(target, "content");

        var queue = AQueue("retry_long_error_queue.json");
        queue.DeleteIf = (_, _) => throw new IOException(new string('e', 10_000));
        queue.Record(target, "locked");

        var first = queue.Sweep();
        Assert.Contains(target, first.Retrying);

        // Read back by a fresh instance, as the next start would.
        var again = AQueue("retry_long_error_queue.json");
        var pending = Assert.Single(again.Pending());
        Assert.Equal(target, pending.Path);
        Assert.True(pending.LastError.Length <= 2048);
        Assert.Equal(1, pending.Attempts);
    }

    [Fact]
    public void AnExpiredDebtOutsideOneHostsAllowlist_ShouldSurviveForTheHostThatOwnsIt()
    {
        // R21-REC06. Kept during retention (R20-REC06) and then abandoned on expiry by the very
        // host that had no business deciding — the covering host never saw it.
        var ours = Directory.CreateDirectory(Path.Combine(TestDir, "ours-expired")).FullName;
        var theirs = Directory.CreateDirectory(Path.Combine(TestDir, "theirs-expired")).FullName;

        var target = Path.Combine(ours, "left_behind_long_ago.txt");
        File.WriteAllText(target, "content");

        var queueFile = CreateTestFilePath("shared_expired_queue.json");
        var covering = new CleanupDebtQueue(queueFile, [ours], Recovery);
        covering.Record(target, "locked");

        var narrow = new CleanupDebtQueue(queueFile, [theirs], Recovery)
        {
            Now = () => DateTimeOffset.UtcNow + CleanupDebtQueue.DefaultRetention + TimeSpan.FromHours(1)
        };

        var narrowSweep = narrow.Sweep();
        Assert.Contains(target, narrowSweep.Refused);
        Assert.DoesNotContain(target, narrowSweep.Abandoned);

        Assert.Contains(target, covering.Pending().Select(d => d.Path));
    }
}
