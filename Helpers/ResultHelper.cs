using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Results;

namespace AsposeMcpServer.Helpers;

/// <summary>
///     Helper class to convert handler results to structured results with output info.
/// </summary>
public static class ResultHelper
{
    /// <summary>
    ///     Finalizes a handler result by wrapping it with output info from the document context.
    /// </summary>
    /// <typeparam name="TDoc">The document context type.</typeparam>
    /// <typeparam name="TResult">The handler result type.</typeparam>
    /// <param name="handlerResult">The handler result object.</param>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">Optional explicit output path override.</param>
    /// <returns>A FinalizedResult containing the data and output info.</returns>
    public static FinalizedResult<TResult> FinalizeResult<TDoc, TResult>(
        TResult handlerResult,
        DocumentContext<TDoc> ctx,
        string? outputPath = null)
        where TDoc : class
    {
        var effectiveOutputPath = outputPath ?? ctx.SourcePath;
        return FinalizeResult(
            handlerResult,
            ctx.IsSession ? null : effectiveOutputPath,
            ctx.IsSession ? ctx.SessionId : null);
    }

    /// <summary>
    ///     Finalizes a handler result by wrapping it with output info.
    /// </summary>
    /// <typeparam name="TResult">The handler result type.</typeparam>
    /// <param name="handlerResult">The handler result object.</param>
    /// <param name="outputPath">Output path for file mode operations.</param>
    /// <param name="sessionId">Session ID for session mode operations.</param>
    /// <returns>A FinalizedResult containing the data and output info.</returns>
    public static FinalizedResult<TResult> FinalizeResult<TResult>(
        TResult handlerResult,
        string? outputPath,
        string? sessionId)
    {
        return new FinalizedResult<TResult>
        {
            Data = handlerResult,
            Output = new OutputInfo
            {
                Path = outputPath,
                SessionId = sessionId,
                IsSession = sessionId != null
            }
        };
    }
}
