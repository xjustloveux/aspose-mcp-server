using System.Diagnostics;
using System.IO.Hashing;
using System.Text.Json;

namespace AsposeMcpServer.Core.Extension.Transport;

/// <summary>
///     Stdin-based transport for extensions.
///     Sends metadata as JSON followed by binary data via stdin.
/// </summary>
public class StdinTransport : IExtensionTransport
{
    /// <summary>
    ///     Default timeout in milliseconds for stdin write operations.
    /// </summary>
    private const int DefaultWriteTimeoutMs = 30000;

    /// <summary>
    ///     Default maximum data size in bytes (100 MB).
    /// </summary>
    private const long DefaultMaxDataSize = 100 * 1024 * 1024;

    /// <summary>
    ///     Logger instance for diagnostic output.
    /// </summary>
    private readonly ILogger<StdinTransport>? _logger;

    /// <summary>
    ///     Maximum data size in bytes.
    /// </summary>
    private readonly long _maxDataSize;

    /// <summary>
    ///     Timeout in milliseconds for write operations.
    /// </summary>
    private readonly int _writeTimeoutMs;

    /// <summary>
    ///     Initializes a new instance of the <see cref="StdinTransport" /> class.
    /// </summary>
    /// <param name="logger">Optional logger instance.</param>
    /// <param name="writeTimeoutMs">Timeout in milliseconds for write operations.</param>
    /// <param name="maxDataSize">Maximum data size in bytes. Defaults to 100 MB.</param>
    public StdinTransport(
        ILogger<StdinTransport>? logger = null,
        int writeTimeoutMs = DefaultWriteTimeoutMs,
        long maxDataSize = DefaultMaxDataSize)
    {
        _logger = logger;
        _writeTimeoutMs = writeTimeoutMs;
        _maxDataSize = maxDataSize;
    }

    /// <inheritdoc />
    public string Mode => "stdin";

    /// <summary>
    ///     Sends data to the extension process via stdin.
    /// </summary>
    /// <param name="process">The extension process to send data to.</param>
    /// <param name="data">The binary data to send.</param>
    /// <param name="metadata">Metadata about the snapshot being sent.</param>
    /// <param name="cancellationToken">
    ///     Cancellation token for the operation. Note: This token is combined with an internal
    ///     timeout of <see cref="_writeTimeoutMs" /> milliseconds. The operation will be cancelled
    ///     if either the provided token is cancelled or the timeout elapses.
    /// </param>
    /// <returns>
    ///     <c>true</c> if the data was sent successfully; otherwise, <c>false</c>.
    /// </returns>
    public async Task<bool> SendAsync(
        Process process,
        byte[] data,
        ExtensionMetadata metadata,
        CancellationToken cancellationToken = default)
    {
        if (process.HasExited)
            return false;

        if (data.Length > _maxDataSize)
        {
            _logger?.LogWarning(
                "Data size ({Size} bytes) exceeds maximum allowed size ({MaxSize} bytes) for session {SessionId}",
                data.Length, _maxDataSize, metadata.SessionId);
            return false;
        }

        using var timeoutCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        timeoutCts.CancelAfter(_writeTimeoutMs);

        try
        {
            if (!IsStdinAvailable(process))
            {
                _logger?.LogWarning(
                    "Stdin is not available for process (closed or redirected) for session {SessionId}",
                    metadata.SessionId);
                return false;
            }

            metadata.TransportMode = Mode;
            metadata.DataSize = data.Length;
            metadata.Checksum = Crc32.HashToUInt32(data);

            var json = JsonSerializer.Serialize(metadata);
            await process.StandardInput.WriteLineAsync(json.AsMemory(), timeoutCts.Token);

            if (process.HasExited)
            {
                _logger?.LogWarning(
                    "Process exited after metadata write for session {SessionId}",
                    metadata.SessionId);
                return false;
            }

            var lengthBytes = BitConverter.GetBytes((long)data.Length);
            if (!BitConverter.IsLittleEndian)
                Array.Reverse(lengthBytes);

            await process.StandardInput.BaseStream.WriteAsync(lengthBytes, timeoutCts.Token);

            if (process.HasExited)
            {
                _logger?.LogWarning(
                    "Process exited after length write for session {SessionId}",
                    metadata.SessionId);
                return false;
            }

            await process.StandardInput.BaseStream.WriteAsync(data, timeoutCts.Token);
            await process.StandardInput.BaseStream.FlushAsync(timeoutCts.Token);

            return true;
        }
        catch (OperationCanceledException ex) when (!cancellationToken.IsCancellationRequested)
        {
            _logger?.LogWarning(ex,
                "Stdin write timed out after {Timeout}ms for session {SessionId}",
                _writeTimeoutMs, metadata.SessionId);
            return false;
        }
        catch (IOException ioEx) when (IsStdinClosedError(ioEx))
        {
            _logger?.LogWarning(ioEx,
                "Stdin was closed during write for session {SessionId}. Process may have terminated.",
                metadata.SessionId);
            return false;
        }
        catch (ObjectDisposedException ex)
        {
            _logger?.LogWarning(ex,
                "Stdin stream was disposed for session {SessionId}. Process may have terminated.",
                metadata.SessionId);
            return false;
        }
        catch (Exception ex)
        {
            _logger?.LogWarning(ex,
                "Failed to send snapshot via stdin transport for session {SessionId}",
                metadata.SessionId);
            return false;
        }
    }

    /// <inheritdoc />
    public void Cleanup(ExtensionMetadata metadata)
    {
    }

    /// <summary>
    ///     Checks if stdin is available for writing.
    /// </summary>
    /// <param name="process">The process to check.</param>
    /// <returns>True if stdin is available.</returns>
    /// <remarks>
    ///     Accessing StandardInput can throw if:
    ///     - Process has exited (InvalidOperationException)
    ///     - Stdin wasn't redirected (InvalidOperationException)
    ///     - Process object is disposed (ObjectDisposedException)
    ///     All these cases mean stdin is not available.
    /// </remarks>
    private static bool IsStdinAvailable(Process process)
    {
        try
        {
            if (process.HasExited)
                return false;

            var stdin = process.StandardInput;
            return stdin.BaseStream.CanWrite;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    ///     Checks if an IOException indicates a closed stdin pipe.
    /// </summary>
    /// <param name="ex">The exception to check.</param>
    /// <returns>True if the error indicates stdin was closed.</returns>
    /// <remarks>
    ///     HResult codes checked:
    ///     <list type="bullet">
    ///         <item>232 (0xE8): ERROR_NO_DATA - The pipe is being closed (Windows)</item>
    ///         <item>109 (0x6D): ERROR_BROKEN_PIPE - The pipe has been ended (Windows)</item>
    ///     </list>
    ///     Also checks message content as fallback for cross-platform compatibility.
    /// </remarks>
    private static bool IsStdinClosedError(IOException ex)
    {
        const int ERROR_NO_DATA = 232;
        const int ERROR_BROKEN_PIPE = 109;

        var hResult = ex.HResult & 0xFFFF;
        return hResult == ERROR_NO_DATA || hResult == ERROR_BROKEN_PIPE ||
               ex.Message.Contains("pipe", StringComparison.OrdinalIgnoreCase) ||
               ex.Message.Contains("broken", StringComparison.OrdinalIgnoreCase) ||
               ex.Message.Contains("closed", StringComparison.OrdinalIgnoreCase);
    }
}
