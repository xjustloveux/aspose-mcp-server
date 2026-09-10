using System.Diagnostics.CodeAnalysis;

namespace AsposeMcpServer.Helpers;

/// <summary>
///     Serialises this process's use of Aspose.Slides.
///     <para>
///         Measured twice, in two different tests, with the same underlying failure:
///         <c>InvalidOperationException: Nullable object must have a value</c> raised from inside
///         Aspose's own code — once from an obfuscated licence/metering frame entered from
///         <c>Shapes.AddAutoShape</c>, once from <c>Aspose.Slides.Presentation..ctor(Stream)</c>
///         with a readable stack. Nothing in this repository is on either stack. Reading or
///         building a presentation while another thread is inside the library is enough to hit it
///         (§19.10.1).
///     </para>
///     <para>
///         <see cref="Core.Session.DocumentSession" /> already holds a semaphore per session, so
///         one session's operations never overlap. Two sessions are a different matter, and so is
///         a file-path operation that never opens a session — those had nothing between them and
///         the library. This gate closes that, at the cost of PowerPoint requests not running
///         concurrently with each other. Word, Excel and PDF are untouched.
///     </para>
///     <para>
///         A mitigation, not a fix: the defect is in the library. It is here because the failure is
///         reachable from the server, not only from the test suite.
///     </para>
/// </summary>
public static class SlidesGate
{
    private static readonly SemaphoreSlim Gate = new(1, 1);

    /// <summary>
    ///     How deep the current thread already is inside the gate. Operations nest — a handler
    ///     opens a second presentation to merge or to read a theme while its own is held — and a
    ///     semaphore is not reentrant, so without this the second entry would wait on the first
    ///     forever.
    /// </summary>
    [ThreadStatic] private static int _depth;

    /// <summary>
    ///     Waits until this thread may use Aspose.Slides, and releases on dispose.
    ///     <para>
    ///         The returned scope must be disposed on the thread that took it. The nesting depth
    ///         is thread-static, so disposing elsewhere would decrement a thread that never
    ///         entered and leave that one unable to nest. Callers hold the gate across a
    ///         synchronous operation, which satisfies this; the scope checks rather than trusts,
    ///         so a future <c>await</c> inside a hold fails loudly instead of corrupting the count
    ///         (R9-S01).
    ///     </para>
    /// </summary>
    /// <returns>A scope to dispose, on this thread, when the work is finished.</returns>
    public static IDisposable Enter()
    {
        if (_depth++ > 0) return new Scope(false);

        Gate.Wait();
        return new Scope(true);
    }

    /// <summary>One thread's hold on the gate.</summary>
    /// <param name="owns">Whether this scope is the outermost one and therefore releases.</param>
    private sealed class Scope(bool owns) : IDisposable
    {
        private readonly int _ownerThreadId = Environment.CurrentManagedThreadId;
        private bool _released;

        /// <inheritdoc />
        /// <exception cref="InvalidOperationException">
        ///     Thrown when the scope is disposed on a thread other than the one that took it.
        /// </exception>
        [SuppressMessage("Major Code Smell", "S3877:Exceptions should not be thrown from unexpected methods",
            Justification =
                "Cross-thread disposal would corrupt the thread-static nesting counter, so this IDisposable contract deliberately fails before releasing the wrong thread's hold.")]
        [SuppressMessage("Major Code Smell", "S2696:Instance members should not write to static fields",
            Justification =
                "Each scope releases exactly one acquisition from the process-wide gate and its thread-static nesting counter.")]
        public void Dispose()
        {
            if (_released) return;

            if (Environment.CurrentManagedThreadId != _ownerThreadId)
                throw new InvalidOperationException(
                    "A SlidesGate scope was released on a different thread from the one that took "
                    + "it. The gate counts nesting per thread, so the hold must begin and end on "
                    + "one thread — do not await inside it.");

            _released = true;

            _depth--;
            if (owns) Gate.Release();
        }
    }
}
