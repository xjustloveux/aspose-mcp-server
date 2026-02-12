using AsposeMcpServer.Results.Extension;

namespace AsposeMcpServer.Core.Extension;

/// <summary>
///     Result of a binding operation.
/// </summary>
public class BindingResult
{
    private BindingResult(bool isSuccess, SessionBindingInfo? binding, ExtensionErrorCode errorCode, string? error)
    {
        IsSuccess = isSuccess;
        Binding = binding;
        ErrorCode = errorCode;
        Error = error;
    }

    /// <summary>
    ///     Gets whether the operation was successful.
    /// </summary>
    public bool IsSuccess { get; }

    /// <summary>
    ///     Gets the binding information if successful.
    /// </summary>
    public SessionBindingInfo? Binding { get; }

    /// <summary>
    ///     Gets the error code if failed.
    /// </summary>
    public ExtensionErrorCode ErrorCode { get; }

    /// <summary>
    ///     Gets the error message if failed.
    /// </summary>
    public string? Error { get; }

    /// <summary>
    ///     Creates a successful result.
    /// </summary>
    /// <param name="binding">The binding information.</param>
    /// <returns>A successful binding result.</returns>
    public static BindingResult Success(SessionBindingInfo binding)
    {
        return new BindingResult(true, binding, ExtensionErrorCode.None, null);
    }

    /// <summary>
    ///     Creates a failure result with the specified error code.
    /// </summary>
    /// <param name="errorCode">The error code.</param>
    /// <param name="error">The error message.</param>
    /// <returns>A failed binding result.</returns>
    public static BindingResult Failure(ExtensionErrorCode errorCode, string error)
    {
        return new BindingResult(false, null, errorCode, error);
    }

    /// <summary>
    ///     Creates a failure result with internal error code.
    /// </summary>
    /// <param name="error">The error message.</param>
    /// <returns>A failed binding result.</returns>
    public static BindingResult Failure(string error)
    {
        return new BindingResult(false, null, ExtensionErrorCode.InternalError, error);
    }
}
