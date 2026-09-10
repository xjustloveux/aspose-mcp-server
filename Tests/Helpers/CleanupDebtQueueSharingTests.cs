using System.Reflection;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Helpers;

/// <summary>
///     §23.5 item 3: two queues on the same file are two views of one record, and have to be
///     serialised as such.
///     <para>
///         Each instance held its own lock, so two hosts configured with the same temp directory
///         could read, sweep and rewrite the same JSON concurrently — and a write built from a
///         stale read drops whatever the other one recorded in between. The lock is now keyed on
///         the queue's canonical path.
///     </para>
///     <para>
///         Queues on different paths must still not wait on each other, or a busy queue would
///         stall an unrelated one. That is the control.
///     </para>
/// </summary>
public class CleanupDebtQueueSharingTests : TestBase
{
    /// <summary>A file for a debt to point at.</summary>
    /// <param name="name">File name.</param>
    /// <returns>The path.</returns>
    private string AFile(string name)
    {
        var path = CreateTestFilePath(name);
        File.WriteAllText(path, "content");
        return path;
    }

    [Fact]
    public void TwoQueuesOnOneFile_ShouldNotLoseDebtsRecordedConcurrently()
    {
        // Interleaved read-modify-write is what loses records: each instance reads the list,
        // adds one entry and writes the whole thing back.
        var queueFile = CreateTestFilePath("shared_queue.json");
        var first = new CleanupDebtQueue(queueFile, [TestDir], Recovery, CleanupDebtQueue.DefaultRetention);
        var second = new CleanupDebtQueue(queueFile, [TestDir], Recovery, CleanupDebtQueue.DefaultRetention);

        const int each = 40;
        var paths = Enumerable.Range(0, each * 2)
            .Select(i => AFile($"shared_target_{i}.txt"))
            .ToList();

        var writers = new[]
        {
            new Thread(() =>
            {
                for (var i = 0; i < each; i++) first.Record(paths[i], "locked");
            }) { IsBackground = true },
            new Thread(() =>
            {
                for (var i = each; i < each * 2; i++) second.Record(paths[i], "locked");
            }) { IsBackground = true }
        };

        foreach (var writer in writers) writer.Start();
        foreach (var writer in writers) Assert.True(writer.Join(TimeSpan.FromSeconds(30)));

        var recorded = first.Pending().Select(debt => debt.Path).ToHashSet(StringComparer.OrdinalIgnoreCase);

        var lost = paths.Where(path => !recorded.Contains(path)).ToList();
        Assert.True(lost.Count == 0,
            $"{lost.Count} of {paths.Count} debts were lost to concurrent writes on one queue file");
    }

    [Fact]
    public void TwoQueuesOnDifferentFiles_ShouldNotWaitOnEachOther()
    {
        // The control: keying the lock on the path must not turn into one global lock, or an
        // unrelated queue would stall behind a busy one.
        var busy = new CleanupDebtQueue(CreateTestFilePath("busy_queue.json"), [TestDir], Recovery);
        var other = new CleanupDebtQueue(CreateTestFilePath("other_queue.json"), [TestDir], Recovery);

        var target = AFile("busy_target.txt");
        var release = new ManualResetEventSlim(false);
        var holding = new ManualResetEventSlim(false);

        // Hold the busy queue inside its own lock by making its delete block.
        busy.DeleteIf = (_, _) =>
        {
            holding.Set();
            release.Wait(TimeSpan.FromSeconds(20));
            return true;
        };
        busy.Record(target, "locked");

        var sweeper = new Thread(() => busy.Sweep()) { IsBackground = true };
        sweeper.Start();

        try
        {
            Assert.True(holding.Wait(TimeSpan.FromSeconds(10)),
                "the busy queue never reached its delete");

            var finished = new ManualResetEventSlim(false);
            var unrelated = new Thread(() =>
            {
                other.Record(AFile("other_target.txt"), "locked");
                finished.Set();
            }) { IsBackground = true };
            unrelated.Start();

            Assert.True(finished.Wait(TimeSpan.FromSeconds(5)),
                "a queue on a different file waited for an unrelated one");
            unrelated.Join(TimeSpan.FromSeconds(5));
        }
        finally
        {
            release.Set();
            sweeper.Join(TimeSpan.FromSeconds(20));
            release.Dispose();
            holding.Dispose();
        }
    }

    [Fact]
    public void TheGateRegistry_ShouldNotGrowForEveryPathAProcessEverUses()
    {
        // R13-F03: it held one lock object and one key per distinct canonical path, strongly, for
        // the life of the process, with nothing to remove them. A queue holds its own gate for as
        // long as it exists, so an entry whose target has been collected is one nothing is using.
        var before = CleanupDebtQueue.TrackedGates;

        for (var i = 0; i < 400; i++)
            // Constructed and dropped: nothing holds the queue, so nothing holds its gate.
            _ = new CleanupDebtQueue(CreateTestFilePath($"gate_growth_{i}.json"), [TestDir], Recovery);

        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();

        // One more construction is what triggers the sweep.
        _ = new CleanupDebtQueue(CreateTestFilePath("gate_growth_final.json"), [TestDir], Recovery);

        Assert.True(CleanupDebtQueue.TrackedGates < before + 400,
            $"the registry grew to {CleanupDebtQueue.TrackedGates} from {before}, which is one "
            + "entry for every path this process has ever used");
    }

    [Fact]
    public void TwoQueuesOnOneFile_ShouldStillShareTheirGateWhileBothExist()
    {
        // The property the weak registry must not break: while both queues are alive their gate is
        // the same object, so a sweep can never hand the second one a lock the first is not using.
        var path = CreateTestFilePath("gate_shared.json");

        var first = new CleanupDebtQueue(path, [TestDir], Recovery);
        var second = new CleanupDebtQueue(path, [TestDir], Recovery);

        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();

        var third = new CleanupDebtQueue(path, [TestDir], Recovery);

        Assert.Same(GateOf(first), GateOf(second));
        Assert.Same(GateOf(first), GateOf(third));
    }

    /// <summary>The lock a queue was handed, which the registry keeps only weakly.</summary>
    /// <param name="queue">The queue to ask.</param>
    /// <returns>Its gate object.</returns>
    private static object GateOf(CleanupDebtQueue queue)
    {
        var field = typeof(CleanupDebtQueue).GetField("_gate",
            BindingFlags.Instance | BindingFlags.NonPublic);
        Assert.NotNull(field);
        var gate = field.GetValue(queue);
        Assert.NotNull(gate);
        return gate;
    }
}
