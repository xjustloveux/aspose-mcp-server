namespace AsposeMcpServer.Core.Handlers;

/// <summary>
///     Defines the contract for operation handlers in the Aspose MCP Server.
///     Each handler is responsible for a single operation type (e.g., add, delete, replace).
///     This interface enables clean separation between MCP tool definitions and business logic.
/// </summary>
/// <typeparam name="TContext">
///     The document context type (e.g., Aspose.Words.Document, Aspose.Slides.Presentation).
/// </typeparam>
/// <remarks>
///     <para>
///         Handlers should be stateless and focus solely on executing the operation.
///         Document loading/saving is handled by the Tool layer.
///     </para>
///     <para>
///         For more information about the Aspose MCP Server architecture, see:
///         https://github.com/aspose/aspose-mcp-server
///     </para>
/// </remarks>
public interface IOperationHandler<TContext> where TContext : class
{
    /// <summary>
    ///     Gets the operation name that this handler processes.
    ///     This value is matched against the 'operation' parameter (case-insensitive).
    /// </summary>
    /// <example>
    ///     Examples: "add", "delete", "replace", "search", "create", "get"
    /// </example>
    string Operation { get; }

    /// <summary>
    ///     Executes the operation on the provided document context.
    /// </summary>
    /// <param name="context">The document context containing the document and metadata.</param>
    /// <param name="parameters">The operation parameters from the MCP request.</param>
    /// <returns>The result object describing the operation outcome.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or invalid.</exception>
    object Execute(OperationContext<TContext> context, OperationParameters parameters);
}
