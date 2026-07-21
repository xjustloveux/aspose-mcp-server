using System.Diagnostics.CodeAnalysis;

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
    ///     Tracks whether this session has been disposed (0 = not disposed, 1 = disposed)
    /// </summary>
    private int _disposed;

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
        WaitForActiveUsersToDrain();

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
            if (Document is IDisposable disposable) disposable.Dispose();
        }
        finally
        {
            if (lockTaken) _lock.Release();
        }
    }

    /// <summary>
    ///     Waits until no usage scopes remain active, up to the given timeout. Close paths call this
    ///     before saving/disposing so they cannot observe (or destroy) a document mid-operation.
    /// </summary>
    /// <param name="timeoutMs">Maximum time to wait in milliseconds.</param>
    /// <returns><c>true</c> when the session drained; <c>false</c> on timeout.</returns>
    internal bool WaitForActiveUsersToDrain(int timeoutMs = DrainTimeoutMs)
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
    public IDisposable AcquireUsage()
    {
        Interlocked.Increment(ref _activeUsers);
        if (IsDisposed)
        {
            Interlocked.Decrement(ref _activeUsers);
            throw new ObjectDisposedException(nameof(DocumentSession),
                $"Session {SessionId} has been disposed");
        }

        return new UsageScope(this);
    }

    /// <summary>
    ///     Releases a usage scope, decrementing the active user count
    /// </summary>
    private void ReleaseUsage()
    {
        Interlocked.Decrement(ref _activeUsers);
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
