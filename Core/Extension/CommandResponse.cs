namespace AsposeMcpServer.Core.Extension;

/// <summary>
///     Response from a command sent to an extension.
/// </summary>
public class CommandResponse
{
    /// <summary>
    ///     Whether the command was successful.
    /// </summary>
    public bool IsSuccess { get; init; }

    /// <summary>
    ///     Command ID that this response corresponds to.
    /// </summary>
    public string? CommandId { get; init; }

    /// <summary>
    ///     Error message if the command failed.
    /// </summary>
    public string? Error { get; init; }

    /// <summary>
    ///     Result data from the extension.
    /// </summary>
    public Dictionary<string, object>? Result { get; init; }

    /// <summary>
    ///     Creates a successful command response.
    /// </summary>
    /// <param name="commandId">The command ID.</param>
    /// <param name="result">Optional result data.</param>
    /// <returns>A successful CommandResponse.</returns>
    public static CommandResponse Success(string commandId, Dictionary<string, object>? result = null)
    {
        return new CommandResponse
        {
            IsSuccess = true,
            CommandId = commandId,
            Result = result
        };
    }

    /// <summary>
    ///     Creates a failed command response.
    /// </summary>
    /// <param name="error">Error message.</param>
    /// <param name="commandId">Optional command ID.</param>
    /// <returns>A failed CommandResponse.</returns>
    public static CommandResponse Failure(string error, string? commandId = null)
    {
        return new CommandResponse
        {
            IsSuccess = false,
            CommandId = commandId,
            Error = error
        };
    }
}
