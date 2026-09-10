using AsposeMcpServer.Helpers;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Helpers;

/// <summary>
///     The gate that keeps this process's Aspose.Slides work to one thread at a time.
///     <para>
///         It exists because the library raised <c>Nullable object must have a value</c> from its
///         own code — twice, in two different tests, once while constructing a presentation and
///         once while adding a shape — whenever another thread was inside it. The session lock is
///         per session and file-path operations open no session at all, so before this there was
///         nothing between two requests and the library (§19.10.1).
///     </para>
///     <para>
///         These tests drive raw <see cref="Thread" />s rather than tasks. The gate tracks nesting
///         in a <c>[ThreadStatic]</c> field, so a hold must begin and end on the same thread; an
///         async test could resume its continuation elsewhere and decrement a thread that never
///         entered.
///     </para>
/// </summary>
public class SlidesGateTests : TestBase
{
    /// <summary>
    ///     Runs <paramref name="body" /> on the given number of dedicated threads and waits for
    ///     them all to finish.
    /// </summary>
    /// <param name="count">How many threads to start.</param>
    /// <param name="timeout">How long to allow them all together.</param>
    /// <param name="body">The work each thread performs.</param>
    /// <returns>Whether every thread finished within the timeout.</returns>
    private static bool RunOnThreads(int count, TimeSpan timeout, Action body)
    {
        var threads = Enumerable.Range(0, count)
            .Select(_ => new Thread(() => body()) { IsBackground = true })
            .ToArray();

        foreach (var thread in threads) thread.Start();

        var deadline = DateTime.UtcNow + timeout;
        foreach (var thread in threads)
        {
            var remaining = deadline - DateTime.UtcNow;
            if (remaining <= TimeSpan.Zero || !thread.Join(remaining)) return false;
        }

        return true;
    }

    [Fact]
    public void TwoThreads_ShouldNotBeInsideTheGateAtTheSameTime()
    {
        var inside = 0;
        var overlapped = false;

        var finished = RunOnThreads(8, TimeSpan.FromSeconds(30), () =>
        {
            for (var round = 0; round < 50; round++)
            {
                using var gate = SlidesGate.Enter();

                if (Interlocked.Increment(ref inside) != 1) overlapped = true;
                Thread.Sleep(1);
                Interlocked.Decrement(ref inside);
            }
        });

        Assert.True(finished, "the gate deadlocked or was far slower than expected");
        Assert.False(overlapped, "two threads were inside the gate at once");
    }

    [Fact]
    public void TheSameThreadEnteringTwice_ShouldNotDeadlock()
    {
        // Operations nest: merging opens a second presentation while holding its own, and applying
        // a theme opens the theme. A plain semaphore would have that wait on itself forever.
        using var outer = SlidesGate.Enter();
        using var inner = SlidesGate.Enter();

        Assert.NotNull(inner);
    }

    [Fact]
    public void ReleasingTheInnerHold_ShouldNotOpenTheGateToAnotherThread()
    {
        using var entered = new ManualResetEventSlim(false);
        Thread other;

        using (SlidesGate.Enter())
        {
            using (SlidesGate.Enter())
            {
                // inner scope ends here
            }

            other = new Thread(() =>
            {
                using var gate = SlidesGate.Enter();
                entered.Set();
            }) { IsBackground = true };
            other.Start();

            Assert.False(entered.Wait(TimeSpan.FromMilliseconds(250)),
                "another thread got in while this one still held the outer scope");
        }

        Assert.True(other.Join(TimeSpan.FromSeconds(10)),
            "the waiting thread never got in after the outer scope ended");
        Assert.True(entered.IsSet);
    }

    [Fact]
    public void DisposingTwice_ShouldReleaseOnlyOnce()
    {
        // A double release would raise the semaphore's count and let two threads in afterwards.
        var scope = SlidesGate.Enter();
        scope.Dispose();
        scope.Dispose();

        var inside = 0;
        var overlapped = false;

        var finished = RunOnThreads(4, TimeSpan.FromSeconds(10), () =>
        {
            using var gate = SlidesGate.Enter();

            if (Interlocked.Increment(ref inside) != 1) overlapped = true;
            Thread.Sleep(5);
            Interlocked.Decrement(ref inside);
        });

        Assert.True(finished, "the gate deadlocked after a double dispose");
        Assert.False(overlapped, "a double dispose released the gate more than once");
    }
}
