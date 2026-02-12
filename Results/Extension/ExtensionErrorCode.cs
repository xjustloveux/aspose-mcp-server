namespace AsposeMcpServer.Results.Extension;

/// <summary>
///     Error codes for extension operations.
/// </summary>
public enum ExtensionErrorCode
{
    /// <summary>
    ///     No error.
    /// </summary>
    None = 0,

    /// <summary>
    ///     Invalid parameter provided.
    /// </summary>
    InvalidParameter = 1,

    /// <summary>
    ///     Session not found.
    /// </summary>
    SessionNotFound = 2,

    /// <summary>
    ///     Extension not found.
    /// </summary>
    ExtensionNotFound = 3,

    /// <summary>
    ///     Extension is unavailable.
    /// </summary>
    ExtensionUnavailable = 4,

    /// <summary>
    ///     Binding not found.
    /// </summary>
    BindingNotFound = 5,

    /// <summary>
    ///     Output format not supported.
    /// </summary>
    FormatNotSupported = 6,

    /// <summary>
    ///     Conversion failed.
    /// </summary>
    ConversionFailed = 7,

    /// <summary>
    ///     Extension system is disabled.
    /// </summary>
    ExtensionDisabled = 8,

    /// <summary>
    ///     Command execution failed.
    /// </summary>
    CommandFailed = 9,

    /// <summary>
    ///     Command timed out.
    /// </summary>
    CommandTimeout = 10,

    /// <summary>
    ///     Invalid JSON payload.
    /// </summary>
    InvalidPayload = 11,

    /// <summary>
    ///     Extension is currently initializing and not yet available.
    /// </summary>
    ExtensionInitializing = 12,

    /// <summary>
    ///     Internal error.
    /// </summary>
    InternalError = 99
}
