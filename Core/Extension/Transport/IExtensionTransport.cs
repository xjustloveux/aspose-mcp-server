using System.Diagnostics;

namespace AsposeMcpServer.Core.Extension.Transport;

/// <summary>
///     Interface for extension transport mechanisms.
/// </summary>
public interface IExtensionTransport
{
    /// <summary>
    ///     Gets the transport mode identifier (e.g., "mmap", "stdin", "file").
    /// </summary>
    string Mode { get; }

    /// <summary>
    ///     Sends data to the extension process.
    /// </summary>
    /// <param name="process">The extension process.</param>
    /// <param name="data">The data to send.</param>
    /// <param name="metadata">The metadata describing the data.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    /// <returns>True if send was successful; otherwise, false.</returns>
    Task<bool> SendAsync(
        Process process,
        byte[] data,
        ExtensionMetadata metadata,
        CancellationToken cancellationToken = default);

    /// <summary>
    ///     Cleans up resources associated with a sent message.
    /// </summary>
    /// <param name="metadata">The metadata of the message to clean up.</param>
    void Cleanup(ExtensionMetadata metadata);
}
