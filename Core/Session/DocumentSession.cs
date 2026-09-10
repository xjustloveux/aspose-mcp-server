using System.Diagnostics.CodeAnalysis;
using Aspose.Slides;
using AsposeMcpServer.Helpers;

namespace AsposeMcpServer.Core.Session;

/// <summary>
///     Represents an open document session in memory
/// </summary>
public sealed class DocumentSession : IDisposable
{
    /// <summary>
    ///     How long close/dispose paths wait for in-flight operations to drain before resources are
    ///     released. Bounded so a stuck operation cannot hang a close forever.
    /// </summary>
    internal const int DrainTimeoutMs = 30_000;

    /// <summary>
    ///     Semaphore for thread-safe document access
    /// </summary>
    private readonly SemaphoreSlim _lock = new(1, 1);

    /// <summary>
    ///     Tracks the number of active users currently holding a usage scope on this session
    /// </summary>
    private int _activeUsers;

    /// <summary>
    ///     1 once this session is being closed for good. Separate from <see cref="_exclusive" />,
    ///     which a save raises and lowers again: a close that borrowed the save barrier could not
    ///     tell "another operation holds this" from "the drain timed out", reported the second,
    ///     skipped its save and disposed the document under a save that had just taken exclusivity
    ///     but had not yet reached its work — so neither of them wrote it (R4-S06).
    /// </summary>
    private int _closing;

    /// <summary>
    ///     1 once a close has handed the document to whoever still holds it — an exclusive holder
    ///     it could not wait out, or the operations still running when it sealed the session. Read
    ///     from those threads, so it is written and read with volatile semantics.
    /// </summary>
    private int _deferredRelease;

    /// <summary>
    ///     1 once some path has taken responsibility for disposing the document. Close, an
    ///     exclusive holder and the last usage scope can all be the last one out; exactly one of
    ///     them may act on it (R4-S06).
    /// </summary>
    private int _disposeClaimed;

    /// <summary>
    ///     Tracks whether this session has been disposed (0 = not disposed, 1 = disposed)
    /// </summary>
    private int _disposed;

    /// <summary>
    ///     1 while a save or close holds the session exclusively. Draining alone was not enough:
    ///     it waited for the count to reach zero, but nothing stopped a new operation from
    ///     acquiring the instant it did, so the caller could still be saving or disposing a
    ///     document another request had just started using (R2-S09).
    /// </summary>
    private int _exclusive;

    /// <summary>
    ///     Backing field for IsDirty with volatile semantics
    /// </summary>
    private int _isDirty;

    /// <summary>
    ///     Backing field for LastAccessedAt with volatile semantics (stores Ticks)
    /// </summary>
    private long _lastAccessedAtTicks;

    /// <summary>
    ///     Creates a new document session
    /// </summary>
    /// <param name="sessionId">Unique session identifier</param>
    /// <param name="path">Original file path</param>
    /// <param name="type">Document type</param>
    /// <param name="document">The Aspose document object</param>
    /// <param name="mode">Access mode (readonly, readwrite)</param>
    public DocumentSession(string sessionId, string path, DocumentType type, object document, string mode)
    {
        SessionId = sessionId;
        Path = path;
        Type = type;
        Document = document;
        Mode = mode;
        OpenedAt = DateTime.UtcNow;
        _lastAccessedAtTicks = OpenedAt.Ticks;
    }

    /// <summary>
    ///     Unique session identifier
    /// </summary>
    public string SessionId { get; }

    /// <summary>
    ///     Original file path
    /// </summary>
    public string Path { get; }

    /// <summary>
    ///     Document type (Word, Excel, PowerPoint, Pdf)
    /// </summary>
    public DocumentType Type { get; }

    /// <summary>
    ///     The Aspose document object
    /// </summary>
    public object Document { get; }

    /// <summary>
    ///     Access mode (readonly, readwrite)
    /// </summary>
    public string Mode { get; }

    /// <summary>
    ///     Whether the document has unsaved changes (thread-safe via Volatile)
    /// </summary>
    public bool IsDirty
    {
        get => Volatile.Read(ref _isDirty) == 1;
        set => Volatile.Write(ref _isDirty, value ? 1 : 0);
    }

    /// <summary>
    ///     When the session was opened
    /// </summary>
    public DateTime OpenedAt { get; }

    /// <summary>
    ///     Last access time (for idle timeout, thread-safe via Volatile)
    /// </summary>
    public DateTime LastAccessedAt
    {
        get => new(Volatile.Read(ref _lastAccessedAtTicks), DateTimeKind.Utc);
        set => Volatile.Write(ref _lastAccessedAtTicks, value.Ticks);
    }

    /// <summary>
    ///     Client identifier (for multi-client scenarios)
    /// </summary>
    public string? ClientId { get; set; }

    /// <summary>
    ///     Session owner identity for isolation
    /// </summary>
    public SessionIdentity Owner { get; init; } = SessionIdentity.GetAnonymous();

    /// <summary>
    ///     Estimated memory usage in bytes
    /// </summary>
    public long EstimatedMemoryBytes { get; set; }

    /// <summary>
    ///     Whether this session has been disposed
    /// </summary>
    public bool IsDisposed => Volatile.Read(ref _disposed) == 1;

    /// <summary>
    ///     Whether this session has active users currently executing operations
    /// </summary>
    public bool HasActiveUsers => Volatile.Read(ref _activeUsers) > 0;

    /// <summary>
    ///     Operations currently holding this session.
    ///     <para>
    ///         Exposed so a test can assert the count never goes below zero. A scope handed out
    ///         after its own increment had been undone drove it negative, and a negative count
    ///         reads as "idle" to the drain loop (R4-S05); <see cref="HasActiveUsers" /> cannot
    ///         tell that apart from a genuinely idle session.
    ///     </para>
    /// </summary>
    internal int ActiveUsers => Volatile.Read(ref _activeUsers);

    /// <summary>
    ///     Whether this session has been sealed for closing and will never serve another operation.
    /// </summary>
    internal bool IsClosing => Volatile.Read(ref _closing) == 1;

    /// <summary>
    ///     Whether a close handed the document to whoever still holds it rather than taking it.
    ///     <para>
    ///         Set when <see cref="BeginClosing" /> returns
    ///         <see cref="SessionSealOutcome.HeldByAnotherOperation" /> or
    ///         <see cref="SessionSealOutcome.ActiveOperationsRemain" />. The session is unregistered
    ///         by then, so nothing else can ever reach it: without this the document would be left
    ///         to the garbage collector. It is deliberately not "the session is closing" — a close
    ///         that goes on to seal the session owns the document itself, and a holder that
    ///         disposed it on the way out would leave that close saving a disposed session.
    ///     </para>
    /// </summary>
    internal bool DeferredReleaseArmed => Volatile.Read(ref _deferredRelease) == 1;

    /// <summary>
    ///     Invoked by <see cref="BeginClosing" /> once the closing barrier is up and before it
    ///     waits for whatever holds the session exclusively.
    ///     <para>
    ///         The window a close has to enter to reproduce R4-S06 is the few instructions between
    ///         a save's successful <see cref="BeginExclusive" /> and the work it then does. A
    ///         fixture cannot land there by lining threads up, so <see cref="OnExclusiveAcquired" />
    ///         holds the save inside the window and this one tells it the close has arrived. Unset
    ///         in production.
    ///     </para>
    /// </summary>
    internal Action? OnClosingBarrierRaised { get; set; }

    /// <summary>
    ///     Invoked when a close has decided it cannot take the document and is about to hand it
    ///     over, before the flag that hands it over is written.
    ///     <para>
    ///         That gap is the whole of R7-S01: whoever finishes there reads an unarmed flag and
    ///         leaves, and the flag is then armed for nobody. A fixture cannot line two threads up
    ///         on a gap this small, so it acts here instead. Unset in production.
    ///     </para>
    /// </summary>
    internal Action? OnCloseAboutToDefer { get; set; }

    /// <summary>
    ///     Invoked by <see cref="BeginExclusive" /> once it has taken the barrier and before it
    ///     returns, which is exactly where a save is most vulnerable to a concurrent close.
    ///     Unset in production, where it costs a null check.
    /// </summary>
    internal Action? OnExclusiveAcquired { get; set; }

    /// <summary>
    ///     Invoked inside <see cref="AcquireUsage" /> after a caller has seen the exclusive
    ///     barrier and backed its increment out, before it decides what to do about it.
    ///     <para>
    ///         The window is a few instructions wide, so a test cannot line two threads up on it
    ///         with a barrier and expect to land inside. This lets a fixture act there every time,
    ///         which is what makes the R4-S05 evidence real rather than a run that happened not to
    ///         collide. Unset in production, where it costs a null check.
    ///     </para>
    /// </summary>
    internal Action? OnExclusiveBarrierObserved { get; set; }

    /// <summary>
    ///     Invoked inside <see cref="AcquireUsage" /> after the first check has passed and before
    ///     the increment, so a fixture can raise the exclusive barrier exactly there.
    ///     <para>
    ///         Together with <see cref="OnExclusiveBarrierObserved" /> this makes the whole R4-S05
    ///         window reachable: the barrier goes up after the caller has been admitted and comes
    ///         down again while it is backing out. Unset in production.
    ///     </para>
    /// </summary>
    internal Action? OnUsageCheckPassed { get; set; }

    /// <summary>
    ///     Whether the last <see cref="Dispose" /> gave up on releasing the document because
    ///     in-flight operations had not finished within <see cref="DrainTimeoutMs" />. Callers use
    ///     it to report that a close completed without reclaiming the document.
    /// </summary>
    public bool DisposedWithActiveUsers { get; private set; }

    /// <summary>
    ///     Fires immediately after the exclusive barrier comes down and before this session decides
    ///     whether it is the last one out. A fixture uses it to run a close's arming inside that
    ///     window.
    /// </summary>
    internal Action? OnExclusiveReleased { get; set; }

    /// <summary>
    ///     Disposes the session and releases all resources including the document.
    ///     Thread-safe: uses Interlocked to prevent double-dispose. Waits (bounded) for in-flight
    ///     operations — usage scopes and lock holders — to finish first, so the document is never
    ///     disposed out from under an operation that is still using it.
    /// </summary>
    [SuppressMessage("Major Code Smell",
        "S2952:Classes should \"Dispose\" of members from the classes' own \"Dispose\" methods",
        Justification = "The SemaphoreSlim is deliberately not disposed: disposing it would leave any " +
                        "concurrently blocked waiter asleep forever (SemaphoreSlim.Dispose does not wake " +
                        "waiters), and an undisposed SemaphoreSlim holds no unmanaged resources when " +
                        "AvailableWaitHandle is never touched. Waiters woken by the Release below observe " +
                        "the disposed flag and throw ObjectDisposedException.")]
    public void Dispose()
    {
        // Atomically set _disposed to 1, return previous value
        // If previous value was already 1, another thread already disposed
        if (Interlocked.Exchange(ref _disposed, 1) == 1)
            return;

        // New work is already rejected (AcquireUsage / ThrowIfDisposed observe _disposed == 1),
        // so this only waits for operations that entered before the close.
        var drained = WaitForActiveUsersToDrain();

        // Serialize with Execute/GetDocument lock holders before the document goes away.
        var lockTaken = false;
        try
        {
            lockTaken = _lock.Wait(DrainTimeoutMs);
        }
        catch (ObjectDisposedException)
        {
            // Unreachable in practice (_lock is never disposed), kept because Wait is documented to throw.
        }

        try
        {
            // Disposing while an operation still holds the document produces a use-after-dispose in
            // that operation, which is worse than leaving the document to the garbage collector.
            // The session is already unregistered, so nothing new can reach it either way.
            DisposedWithActiveUsers = !drained || !lockTaken;
            if (DisposedWithActiveUsers) return;

            // Releasing a presentation re-enters Aspose.Slides, which is not safe to touch from
            // two threads, and this runs on whichever thread closed, disconnected or timed the
            // session out — never in step with a converting or saving one (SlidesGate).
            using var slidesGate = Document is Presentation
                ? SlidesGate.Enter()
                : null;

            if (Document is IDisposable disposable) disposable.Dispose();
        }
        finally
        {
            if (lockTaken) _lock.Release();
        }
    }

    /// <summary>
    ///     Takes responsibility for disposing the document, once.
    ///     <para>
    ///         Three paths can be the last one out — the close itself, an exclusive holder it
    ///         declined to wait for, and the operation whose usage scope brings the count to zero.
    ///         Each asks here first, so the document is released exactly once and by whichever of
    ///         them actually finished last (R4-S06).
    ///     </para>
    /// </summary>
    /// <returns><c>true</c> for the one caller that may dispose the document.</returns>
    internal bool TryClaimDisposeOwnership()
    {
        return Interlocked.Exchange(ref _disposeClaimed, 1) == 0;
    }

    /// <summary>
    ///     Waits until no usage scopes remain active, up to the given timeout. Close paths call this
    ///     before saving/disposing so they cannot observe (or destroy) a document mid-operation.
    /// </summary>
    /// <param name="timeoutMs">Maximum time to wait in milliseconds.</param>
    /// <returns><c>true</c> when the session drained; <c>false</c> on timeout.</returns>
    internal bool WaitForActiveUsersToDrain(int timeoutMs = DrainTimeoutMs)
    {
        return DrainCore(timeoutMs);
    }

    /// <summary>
    ///     Closes the session to new operations and waits for the in-flight ones to finish.
    ///     <para>
    ///         The barrier is what makes the wait meaningful: once it is set, <see cref="AcquireUsage" />
    ///         refuses, so the count can only fall. A caller that keeps using the session afterwards
    ///         (a save, as opposed to a close) must call <see cref="EndExclusive" />.
    ///     </para>
    /// </summary>
    /// <param name="timeoutMs">Maximum time to wait in milliseconds.</param>
    /// <returns><c>true</c> when the session drained and is now held exclusively.</returns>
    internal bool BeginExclusive(int timeoutMs = DrainTimeoutMs)
    {
        // A sealed session is never lent out again. Checked before and after taking the barrier:
        // a close that begins in between waits for this barrier to fall, so a save admitted there
        // has to back out rather than leave the close waiting on work it has already refused.
        if (IsClosing) return false;

        // CompareExchange, not a plain write: two callers racing to save and close would both
        // have written 1 and both been told they held the session, and whichever finished first
        // would then have released the other's exclusivity (R3-S04).
        if (Interlocked.CompareExchange(ref _exclusive, 1, 0) != 0) return false;

        if (IsClosing)
        {
            Volatile.Write(ref _exclusive, 0);
            return false;
        }

        if (DrainCore(timeoutMs))
        {
            OnExclusiveAcquired?.Invoke();
            return true;
        }

        // Nothing is held on a failed drain, so the barrier must come back down or the session
        // would stay permanently closed to new work.
        Volatile.Write(ref _exclusive, 0);
        return false;
    }

    /// <summary>
    ///     Seals the session for closing and takes final ownership of the document.
    ///     <para>
    ///         A close used <see cref="BeginExclusive" />, whose single <c>false</c> meant both
    ///         "someone else holds it" and "the drain timed out". The close reported the second,
    ///         skipped its save and disposed anyway, which failed a save that had taken exclusivity
    ///         moments earlier and had not yet reached its work: neither path wrote the document
    ///         (R4-S06). Closing is a permanent state of its own, so a save in flight is waited for
    ///         rather than raced, and the outcome says which of the two happened.
    ///     </para>
    /// </summary>
    /// <param name="timeoutMs">Maximum total time to wait in milliseconds.</param>
    /// <returns>What the seal achieved.</returns>
    internal SessionSealOutcome BeginClosing(int timeoutMs = DrainTimeoutMs)
    {
        if (Interlocked.Exchange(ref _closing, 1) == 1) return SessionSealOutcome.AlreadyClosing;

        OnClosingBarrierRaised?.Invoke();

        // New operations and new saves are refused from here on, so both of the waits below can
        // only ever count down.
        var deadline = Environment.TickCount64 + timeoutMs;
        var spin = new SpinWait();
        while (Volatile.Read(ref _exclusive) == 1)
        {
            if (Environment.TickCount64 >= deadline)
            {
                ArmDeferredRelease(false);
                return SessionSealOutcome.HeldByAnotherOperation;
            }

            if (spin.NextSpinWillYield) Thread.Sleep(5);
            else spin.SpinOnce();
        }

        // A plain write rather than a CompareExchange: only one caller reaches this line, because
        // only one wins the exchange above and no save can take the barrier while _closing is 1.
        Volatile.Write(ref _exclusive, 1);

        var remaining = (int)Math.Max(0, deadline - Environment.TickCount64);
        if (DrainCore(remaining)) return SessionSealOutcome.Sealed;

        // The operations that were already running keep the document until they finish. The
        // session is unregistered by now, so the last one out has to release it or nothing will
        // (R4-S06).
        ArmDeferredRelease(true);
        return SessionSealOutcome.ActiveOperationsRemain;
    }

    /// <summary>
    ///     Hands the document to whoever still holds it, and settles the case where nobody does.
    ///     <para>
    ///         The flag used to be written only after the close had finished deciding. Whoever was
    ///         last could release in between, read an unarmed flag, and leave: the close then armed
    ///         a flag no one would ever read again, and the document stayed alive with no owner
    ///         (R7-S01). Arming is followed by re-reading the counts, so this close settles the
    ///         case rather than leaving it to an actor that has already gone.
    ///     </para>
    /// </summary>
    /// <param name="closeHoldsExclusive">
    ///     Whether the exclusive barrier is this close's own. When it is not, an exclusive holder
    ///     that has since released also has to be noticed here.
    /// </param>
    private void ArmDeferredRelease(bool closeHoldsExclusive)
    {
        OnCloseAboutToDefer?.Invoke();

        // Interlocked, not a volatile write. This side writes the flag and then reads the counts
        // while the release side writes a count and then reads the flag — a store followed by a
        // load of a different location, which x86-TSO and the memory model both allow to be
        // reordered. Each side could see the other's old value, and then neither would dispose:
        // an armed session with no owner (R8-S03). An exchange carries a full fence, so the write
        // here is visible before the reads below.
        Interlocked.Exchange(ref _deferredRelease, 1);

        if (Volatile.Read(ref _activeUsers) != 0) return;
        if (!closeHoldsExclusive && Volatile.Read(ref _exclusive) != 0) return;

        if (TryClaimDisposeOwnership()) Dispose();
    }

    /// <summary>
    ///     Reopens the session to new operations after a successful <see cref="BeginExclusive" />.
    ///     <para>
    ///         Only a caller whose <see cref="BeginExclusive" /> returned <c>true</c> may call this.
    ///         A caller that was refused and released anyway would tear down the holder's barrier,
    ///         which is what an unconditional release in a <c>finally</c> did (R3-S04).
    ///     </para>
    /// </summary>
    internal void EndExclusive()
    {
        // The same full fence as the arming side, for the same reason (R8-S03).
        Interlocked.Exchange(ref _exclusive, 0);

        OnExclusiveReleased?.Invoke();

        // And the settle happens here rather than at each call site. A close that gave up waiting
        // leaves the document to whoever holds it, and an exclusive holder is one of the three
        // ways to be last out; the other two already settle in ReleaseUsage and in the close
        // itself. Leaving it to the callers meant every future one had to remember (R8-S03).
        if (DeferredReleaseArmed && Volatile.Read(ref _activeUsers) == 0
                                 && TryClaimDisposeOwnership())
            Dispose();
    }

    /// <summary>
    ///     Spins until no usage scopes remain, or the deadline passes.
    /// </summary>
    /// <param name="timeoutMs">Maximum time to wait in milliseconds.</param>
    /// <returns><c>true</c> when the count reached zero.</returns>
    private bool DrainCore(int timeoutMs)
    {
        var spin = new SpinWait();
        var deadline = Environment.TickCount64 + timeoutMs;
        while (Volatile.Read(ref _activeUsers) > 0)
        {
            if (Environment.TickCount64 >= deadline) return false;
            if (spin.NextSpinWillYield) Thread.Sleep(5);
            else spin.SpinOnce();
        }

        return true;
    }

    /// <summary>
    ///     Throws ObjectDisposedException if session is disposed
    /// </summary>
    /// <exception cref="ObjectDisposedException">Thrown when session is disposed</exception>
    private void ThrowIfDisposed()
    {
        if (IsDisposed)
            throw new ObjectDisposedException(nameof(DocumentSession), $"Session {SessionId} has been disposed");
    }

    /// <summary>
    ///     Acquires a usage scope that prevents the session from being cleaned up while in use.
    ///     The caller must dispose the returned scope when the operation is complete.
    /// </summary>
    /// <returns>A disposable usage scope</returns>
    /// <exception cref="ObjectDisposedException">Thrown when session is already disposed</exception>
    /// <exception cref="InvalidOperationException">
    ///     Thrown when the session is being saved or closed and is not accepting new operations.
    /// </exception>
    public IDisposable AcquireUsage()
    {
        ThrowIfExclusive();

        OnUsageCheckPassed?.Invoke();

        Interlocked.Increment(ref _activeUsers);

        // Re-checked after the increment: a caller that passed the first check before the barrier
        // went up backs out here, which is what lets the drain loop reach zero and stay there.
        //
        // Once it has backed out it must throw unconditionally. Re-reading the flag to decide
        // meant an owner that released exclusivity in between left this method returning a scope
        // whose increment had already been undone, so disposing it drove the count below zero and
        // the next drain saw an idle session while an operation was still running (R4-S05).
        if (Volatile.Read(ref _exclusive) == 1 || IsClosing)
        {
            Interlocked.Decrement(ref _activeUsers);
            OnExclusiveBarrierObserved?.Invoke();
            throw new InvalidOperationException(
                $"Session {SessionId} is being saved or closed and is not accepting new operations.");
        }

        if (IsDisposed)
        {
            Interlocked.Decrement(ref _activeUsers);
            throw new ObjectDisposedException(nameof(DocumentSession),
                $"Session {SessionId} has been disposed");
        }

        return new UsageScope(this);
    }

    /// <summary>
    ///     Refuses a new operation while a save or close holds the session.
    /// </summary>
    /// <exception cref="InvalidOperationException">Thrown when the session is held exclusively.</exception>
    private void ThrowIfExclusive()
    {
        if (Volatile.Read(ref _exclusive) == 1 || IsClosing)
            throw new InvalidOperationException(
                $"Session {SessionId} is being saved or closed and cannot accept new operations.");
    }

    /// <summary>
    ///     Releases a usage scope, decrementing the active user count
    /// </summary>
    private void ReleaseUsage()
    {
        var remaining = Interlocked.Decrement(ref _activeUsers);

        // A close that could not wait for this operation left the document here. Only the caller
        // that brings the count to zero can be the last one out, and only one of the three paths
        // that can be last is allowed to act on it.
        if (remaining == 0 && DeferredReleaseArmed && TryClaimDisposeOwnership())
            Dispose();
    }

    /// <summary>
    ///     Execute a synchronous operation on the document with thread-safety
    /// </summary>
    /// <param name="operation">The operation to execute on the document</param>
    /// <exception cref="ObjectDisposedException">Thrown when session is disposed</exception>
    public void Execute(Action<object> operation)
    {
        ThrowIfDisposed();
        try
        {
            _lock.Wait();
        }
        catch (ObjectDisposedException)
        {
            throw new ObjectDisposedException(nameof(DocumentSession), $"Session {SessionId} has been disposed");
        }

        try
        {
            ThrowIfDisposed();
            LastAccessedAt = DateTime.UtcNow;
            operation(Document);
        }
        finally
        {
            _lock.Release();
        }
    }

    /// <summary>
    ///     Execute an operation on the document with thread-safety
    /// </summary>
    /// <typeparam name="T">Return type of the operation</typeparam>
    /// <param name="operation">The operation to execute on the document</param>
    /// <param name="cancellationToken">Cancellation token</param>
    /// <returns>Result of the operation</returns>
    /// <exception cref="ObjectDisposedException">Thrown when session is disposed</exception>
    public async Task<T> ExecuteAsync<T>(Func<object, T> operation, CancellationToken cancellationToken = default)
    {
        ThrowIfDisposed();
        try
        {
            await _lock.WaitAsync(cancellationToken);
        }
        catch (ObjectDisposedException)
        {
            throw new ObjectDisposedException(nameof(DocumentSession), $"Session {SessionId} has been disposed");
        }

        try
        {
            ThrowIfDisposed(); // Re-check after acquiring lock
            LastAccessedAt = DateTime.UtcNow;
            return operation(Document);
        }
        finally
        {
            _lock.Release();
        }
    }

    /// <summary>
    ///     Execute an async operation on the document with thread-safety
    /// </summary>
    /// <typeparam name="T">Return type of the operation</typeparam>
    /// <param name="operation">The async operation to execute on the document</param>
    /// <param name="cancellationToken">Cancellation token</param>
    /// <returns>Result of the operation</returns>
    /// <exception cref="ObjectDisposedException">Thrown when session is disposed</exception>
    public async Task<T> ExecuteAsync<T>(Func<object, Task<T>> operation, CancellationToken cancellationToken = default)
    {
        ThrowIfDisposed();
        try
        {
            await _lock.WaitAsync(cancellationToken);
        }
        catch (ObjectDisposedException)
        {
            throw new ObjectDisposedException(nameof(DocumentSession), $"Session {SessionId} has been disposed");
        }

        try
        {
            ThrowIfDisposed(); // Re-check after acquiring lock
            LastAccessedAt = DateTime.UtcNow;
            return await operation(Document);
        }
        finally
        {
            _lock.Release();
        }
    }

    /// <summary>
    ///     Get the document as a specific type (thread-safe)
    /// </summary>
    /// <typeparam name="T">Target document type</typeparam>
    /// <returns>Document cast to type T</returns>
    /// <exception cref="InvalidCastException">Thrown when document is not of type T</exception>
    /// <exception cref="ObjectDisposedException">Thrown when session is disposed</exception>
    public T GetDocument<T>() where T : class
    {
        ThrowIfDisposed();
        try
        {
            _lock.Wait();
        }
        catch (ObjectDisposedException)
        {
            throw new ObjectDisposedException(nameof(DocumentSession), $"Session {SessionId} has been disposed");
        }

        try
        {
            ThrowIfDisposed(); // Re-check after acquiring lock
            LastAccessedAt = DateTime.UtcNow;
            return Document as T ?? throw new InvalidCastException($"Document is not of type {typeof(T).Name}");
        }
        finally
        {
            _lock.Release();
        }
    }

    /// <summary>
    ///     Get the document as a specific type asynchronously (thread-safe)
    /// </summary>
    /// <typeparam name="T">Target document type</typeparam>
    /// <param name="cancellationToken">Cancellation token</param>
    /// <returns>Document cast to type T</returns>
    /// <exception cref="InvalidCastException">Thrown when document is not of type T</exception>
    /// <exception cref="ObjectDisposedException">Thrown when session is disposed</exception>
    public async Task<T> GetDocumentAsync<T>(CancellationToken cancellationToken = default) where T : class
    {
        ThrowIfDisposed();
        try
        {
            await _lock.WaitAsync(cancellationToken);
        }
        catch (ObjectDisposedException)
        {
            throw new ObjectDisposedException(nameof(DocumentSession), $"Session {SessionId} has been disposed");
        }

        try
        {
            ThrowIfDisposed(); // Re-check after acquiring lock
            LastAccessedAt = DateTime.UtcNow;
            return Document as T ?? throw new InvalidCastException($"Document is not of type {typeof(T).Name}");
        }
        finally
        {
            _lock.Release();
        }
    }

    /// <summary>
    ///     Represents a usage scope that prevents the session from being cleaned up.
    ///     Disposing this scope releases the usage count.
    /// </summary>
    private sealed class UsageScope : IDisposable
    {
        private DocumentSession? _session;

        public UsageScope(DocumentSession session)
        {
            _session = session;
        }

        public void Dispose()
        {
            var s = Interlocked.Exchange(ref _session, null);
            s?.ReleaseUsage();
        }
    }
}
