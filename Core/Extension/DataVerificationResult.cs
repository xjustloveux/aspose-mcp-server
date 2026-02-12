namespace AsposeMcpServer.Core.Extension;

/// <summary>
///     Result of data verification against metadata.
/// </summary>
public enum DataVerificationResult
{
    /// <summary>
    ///     Data is valid - size and checksum match.
    /// </summary>
    Valid,

    /// <summary>
    ///     Data is null.
    /// </summary>
    NullData,

    /// <summary>
    ///     Data size does not match expected size.
    /// </summary>
    SizeMismatch,

    /// <summary>
    ///     Checksum does not match - data may be corrupted.
    /// </summary>
    ChecksumMismatch
}
