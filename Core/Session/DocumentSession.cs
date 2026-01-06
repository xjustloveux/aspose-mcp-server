namespace AsposeMcpServer.Core.Session;

/// <summary>
///     Represents an open document session in memory
/// </summary>
public class DocumentSession : IDisposable
{
    /// <summary>
    ///     Semaphore for thread-safe document access
    /// </summary>
    private readonly SemaphoreSlim _lock = new(1, 1);

    /// <summary>
    ///     Tracks whether this session has been disposed (0 = not disposed, 1 = disposed)
    /// </summary>
    private int _disposed;

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
        LastAccessedAt = DateTime.UtcNow;
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
    ///     Whether the document has unsaved changes
    /// </summary>
    public bool IsDirty { get; set; }

    /// <summary>
    ///     When the session was opened
    /// </summary>
    public DateTime OpenedAt { get; }

    /// <summary>
    ///     Last access time (for idle timeout)
    /// </summary>
    public DateTime LastAccessedAt { get; set; }

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
    ///     Disposes the session and releases all resources including the document.
    ///     Thread-safe: uses Interlocked to prevent double-dispose.
    /// </summary>
    public void Dispose()
    {
        // Atomically set _disposed to 1, return previous value
        // If previous value was already 1, another thread already disposed
        if (Interlocked.Exchange(ref _disposed, 1) == 1)
            return;

        _lock.Dispose();

        if (Document is IDisposable disposable) disposable.Dispose();
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
}