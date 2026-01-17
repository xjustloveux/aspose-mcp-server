using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;

namespace AsposeMcpServer.Core.Tools;

/// <summary>
///     Base class for properties tools that manage document metadata.
///     Provides common functionality for get/set properties operations.
/// </summary>
/// <typeparam name="TDocument">The document type (e.g., Document, Workbook).</typeparam>
public abstract class PropertiesToolBase<TDocument> where TDocument : class
{
    /// <summary>
    ///     Handler registry for properties operations.
    /// </summary>
    protected readonly HandlerRegistry<TDocument> HandlerRegistry;

    /// <summary>
    ///     Session identity accessor for session isolation support.
    /// </summary>
    protected readonly ISessionIdentityAccessor? IdentityAccessor;

    /// <summary>
    ///     Document session manager for in-memory editing support.
    /// </summary>
    protected readonly DocumentSessionManager? SessionManager;

    /// <summary>
    ///     Initializes a new instance of the properties tool base class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional session identity accessor for session isolation.</param>
    /// <param name="handlerNamespace">The namespace containing the operation handlers.</param>
    protected PropertiesToolBase(
        DocumentSessionManager? sessionManager,
        ISessionIdentityAccessor? identityAccessor,
        string handlerNamespace)
    {
        SessionManager = sessionManager;
        IdentityAccessor = identityAccessor;
        HandlerRegistry = HandlerRegistry<TDocument>.CreateFromNamespace(handlerNamespace);
    }

    /// <summary>
    ///     Executes a properties operation with common workflow.
    /// </summary>
    /// <param name="operation">The operation to perform.</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="path">Document file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="parameters">The operation parameters.</param>
    /// <param name="isReadOnlyOperation">Function to determine if the operation is read-only.</param>
    /// <returns>The operation result.</returns>
    protected string ExecuteOperation(
        string operation,
        string? sessionId,
        string? path,
        string? outputPath,
        OperationParameters parameters,
        Func<string, bool> isReadOnlyOperation)
    {
        using var ctx = DocumentContext<TDocument>.Create(SessionManager, sessionId, path, IdentityAccessor);

        var handler = HandlerRegistry.GetHandler(operation);
        var effectiveOutputPath = outputPath ?? path;

        var operationContext = new OperationContext<TDocument>
        {
            Document = ctx.Document,
            SessionManager = SessionManager,
            IdentityAccessor = IdentityAccessor,
            SessionId = sessionId,
            SourcePath = path,
            OutputPath = effectiveOutputPath
        };

        var result = handler.Execute(operationContext, parameters);

        // Read-only operations return result directly
        if (isReadOnlyOperation(operation))
            return result;

        // Write operations save and return with output message
        if (operationContext.IsModified)
            ctx.Save(effectiveOutputPath);

        if (ctx.IsSession || !operationContext.IsModified)
            return result;

        return $"{result}\n{ctx.GetOutputMessage(effectiveOutputPath)}";
    }
}
