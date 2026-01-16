using System.ComponentModel;
using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for managing PowerPoint sections (add, rename, delete, get).
/// </summary>
[McpServerToolType]
public class PptSectionTool
{
    /// <summary>
    ///     Handler registry for section operations.
    /// </summary>
    private readonly HandlerRegistry<Presentation> _handlerRegistry;

    /// <summary>
    ///     Identity accessor for session isolation.
    /// </summary>
    private readonly ISessionIdentityAccessor? _identityAccessor;

    /// <summary>
    ///     Session manager for document session handling.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="PptSectionTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory editing.</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation.</param>
    public PptSectionTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry =
            HandlerRegistry<Presentation>.CreateFromNamespace("AsposeMcpServer.Handlers.PowerPoint.Section");
    }

    /// <summary>
    ///     Executes a PowerPoint section operation (add, rename, delete, get).
    /// </summary>
    /// <param name="operation">The operation to perform: add, rename, delete, get.</param>
    /// <param name="path">Presentation file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="name">Section name (required for add).</param>
    /// <param name="slideIndex">Start slide index for section (0-based, required for add).</param>
    /// <param name="sectionIndex">Section index (0-based, required for rename/delete).</param>
    /// <param name="newName">New section name (required for rename).</param>
    /// <param name="keepSlides">Keep slides in presentation (optional, for delete, default: true).</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "ppt_section")]
    [Description(@"Manage PowerPoint sections. Supports 4 operations: add, rename, delete, get.

Warning: If outputPath is not provided for add/rename/delete operations, the original file will be overwritten.

Usage examples:
- Add section: ppt_section(operation='add', path='presentation.pptx', name='Section 1', slideIndex=0)
- Rename section: ppt_section(operation='rename', path='presentation.pptx', sectionIndex=0, newName='New Section')
- Delete section: ppt_section(operation='delete', path='presentation.pptx', sectionIndex=0)
- Get sections: ppt_section(operation='get', path='presentation.pptx')")]
    public string Execute(
        [Description("Operation: add, rename, delete, get")]
        string operation,
        [Description("Presentation file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Section name (required for add)")]
        string? name = null,
        [Description("Start slide index for section (0-based, required for add)")]
        int? slideIndex = null,
        [Description("Section index (0-based, required for rename/delete)")]
        int? sectionIndex = null,
        [Description("New section name (required for rename)")]
        string? newName = null,
        [Description("Keep slides in presentation (optional, for delete, default: true)")]
        bool keepSlides = true)
    {
        using var ctx = DocumentContext<Presentation>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters = BuildParameters(operation, name, slideIndex, sectionIndex, newName, keepSlides);

        var handler = _handlerRegistry.GetHandler(operation);

        var operationContext = new OperationContext<Presentation>
        {
            Document = ctx.Document,
            SessionManager = _sessionManager,
            IdentityAccessor = _identityAccessor,
            SessionId = sessionId,
            SourcePath = path,
            OutputPath = outputPath
        };

        var result = handler.Execute(operationContext, parameters);

        if (string.Equals(operation, "get", StringComparison.OrdinalIgnoreCase))
            return result;

        if (operationContext.IsModified)
            ctx.Save(outputPath);

        return $"{result}\n{ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Builds OperationParameters from method parameters using strategy pattern.
    /// </summary>
    private static OperationParameters BuildParameters(
        string operation,
        string? name,
        int? slideIndex,
        int? sectionIndex,
        string? newName,
        bool keepSlides)
    {
        return operation.ToLowerInvariant() switch
        {
            "add" => BuildAddParameters(name, slideIndex),
            "rename" => BuildRenameParameters(sectionIndex, newName),
            "delete" => BuildDeleteParameters(sectionIndex, keepSlides),
            _ => new OperationParameters()
        };
    }

    /// <summary>
    ///     Builds parameters for the add section operation.
    /// </summary>
    /// <param name="name">The section name.</param>
    /// <param name="slideIndex">The start slide index for the section (0-based).</param>
    /// <returns>OperationParameters configured for adding a section.</returns>
    private static OperationParameters BuildAddParameters(string? name, int? slideIndex)
    {
        var parameters = new OperationParameters();
        if (name != null) parameters.Set("name", name);
        if (slideIndex.HasValue) parameters.Set("slideIndex", slideIndex.Value);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the rename section operation.
    /// </summary>
    /// <param name="sectionIndex">The section index (0-based).</param>
    /// <param name="newName">The new section name.</param>
    /// <returns>OperationParameters configured for renaming a section.</returns>
    private static OperationParameters BuildRenameParameters(int? sectionIndex, string? newName)
    {
        var parameters = new OperationParameters();
        if (sectionIndex.HasValue) parameters.Set("sectionIndex", sectionIndex.Value);
        if (newName != null) parameters.Set("newName", newName);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the delete section operation.
    /// </summary>
    /// <param name="sectionIndex">The section index (0-based).</param>
    /// <param name="keepSlides">Whether to keep slides in the presentation.</param>
    /// <returns>OperationParameters configured for deleting a section.</returns>
    private static OperationParameters BuildDeleteParameters(int? sectionIndex, bool keepSlides)
    {
        var parameters = new OperationParameters();
        if (sectionIndex.HasValue) parameters.Set("sectionIndex", sectionIndex.Value);
        parameters.Set("keepSlides", keepSlides);
        return parameters;
    }
}
