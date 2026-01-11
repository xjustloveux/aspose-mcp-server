using AsposeMcpServer.Core.Session;

namespace AsposeMcpServer.Core.Handlers;

/// <summary>
///     Contains the document and session context for handler execution.
///     This class encapsulates all the context information needed by a handler
///     to perform its operation without coupling to MCP-specific concepts.
/// </summary>
/// <typeparam name="TContext">
///     The document type (e.g., Aspose.Words.Document, Aspose.Slides.Presentation).
/// </typeparam>
/// <remarks>
///     <para>
///         OperationContext provides a clean abstraction layer between the MCP Tool
///         interface and the Handler business logic. Tools are responsible for creating
///         the context, while Handlers consume it.
///     </para>
///     <para>
///         For session-based operations, the context includes session management
///         references. For file-based operations, these may be null.
///     </para>
/// </remarks>
public class OperationContext<TContext> where TContext : class
{
    /// <summary>
    ///     Gets the document instance to operate on.
    ///     This is the Aspose document object that the handler will modify or query.
    /// </summary>
    public required TContext Document { get; init; }

    /// <summary>
    ///     Gets the session manager for session-based operations.
    ///     May be null for file-based operations where documents are loaded directly from disk.
    /// </summary>
    public DocumentSessionManager? SessionManager { get; init; }

    /// <summary>
    ///     Gets the session identity accessor for session isolation.
    ///     Used to retrieve the current user/session identity for authorization.
    ///     May be null for file-based operations.
    /// </summary>
    public ISessionIdentityAccessor? IdentityAccessor { get; init; }

    /// <summary>
    ///     Gets the session ID if operating in session mode.
    ///     This is the unique identifier for the in-memory document session.
    /// </summary>
    public string? SessionId { get; init; }

    /// <summary>
    ///     Gets the source file path.
    ///     For file-based operations, this is the path from which the document was loaded.
    /// </summary>
    public string? SourcePath { get; init; }

    /// <summary>
    ///     Gets the output file path for save operations.
    ///     If different from SourcePath, the document will be saved to a new location.
    /// </summary>
    public string? OutputPath { get; init; }

    /// <summary>
    ///     Gets or sets whether the document has been modified by the operation.
    ///     Handlers should set this to true when they make changes to the document.
    ///     The Tool layer uses this flag to determine whether to save the document.
    /// </summary>
    public bool IsModified { get; set; }

    /// <summary>
    ///     Gets or sets the result document for operations that create a new document
    ///     instead of modifying the existing one (e.g., delete_page).
    ///     When set, the Tool layer should save this document instead of the original.
    /// </summary>
    public TContext? ResultDocument { get; set; }
}
