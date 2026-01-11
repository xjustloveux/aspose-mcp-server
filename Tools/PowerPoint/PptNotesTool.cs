using System.ComponentModel;
using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Handlers.PowerPoint.Notes;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for managing PowerPoint notes.
///     Supports: set, get, clear, set_header_footer
///     Note: Notes pages have separate header and footer fields (unlike slides which only have footer).
/// </summary>
[McpServerToolType]
public class PptNotesTool
{
    /// <summary>
    ///     Handler registry for notes operations.
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
    ///     Initializes a new instance of the <see cref="PptNotesTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory editing.</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation.</param>
    public PptNotesTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry = PptNotesHandlerRegistry.Create();
    }

    /// <summary>
    ///     Executes a PowerPoint notes operation (set, get, clear, set_header_footer).
    /// </summary>
    /// <param name="operation">The operation to perform: set, get, clear, set_header_footer.</param>
    /// <param name="path">Presentation file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="slideIndex">Slide index (0-based, required for set, optional for get).</param>
    /// <param name="notes">Notes text content (required for set). Will replace existing notes.</param>
    /// <param name="slideIndices">Slide indices array (optional for clear, if not provided affects all slides).</param>
    /// <param name="headerText">Header text for notes pages (optional for set_header_footer).</param>
    /// <param name="footerText">Footer text for notes pages (optional for set_header_footer).</param>
    /// <param name="dateText">Date/time text for notes pages (optional for set_header_footer).</param>
    /// <param name="showPageNumber">Show page number on notes pages (optional for set_header_footer, default: true).</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "ppt_notes")]
    [Description(@"Manage PowerPoint notes. Supports 4 operations: set, get, clear, set_header_footer.

Note: Notes pages have separate header and footer fields (unlike slides which only have footer).
Warning: 'set' operation will REPLACE existing notes content (format will be reset).
Warning: If outputPath is not provided, the original file will be overwritten.

Usage examples:
- Set notes: ppt_notes(operation='set', path='presentation.pptx', slideIndex=0, notes='Speaker notes')
- Get notes: ppt_notes(operation='get', path='presentation.pptx', slideIndex=0)
- Clear notes: ppt_notes(operation='clear', path='presentation.pptx', slideIndices=[0,1,2])
- Set header/footer: ppt_notes(operation='set_header_footer', path='presentation.pptx', headerText='Header', footerText='Footer')")]
    public string Execute(
        [Description("Operation: set, get, clear, set_header_footer")]
        string operation,
        [Description("Presentation file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Slide index (0-based, required for set, optional for get)")]
        int? slideIndex = null,
        [Description("Notes text content (required for set). Will replace existing notes.")]
        string? notes = null,
        [Description("Slide indices array (optional for clear, if not provided affects all slides)")]
        int[]? slideIndices = null,
        [Description("Header text for notes pages (optional for set_header_footer)")]
        string? headerText = null,
        [Description("Footer text for notes pages (optional for set_header_footer)")]
        string? footerText = null,
        [Description("Date/time text for notes pages (optional for set_header_footer)")]
        string? dateText = null,
        [Description("Show page number on notes pages (optional for set_header_footer, default: true)")]
        bool showPageNumber = true)
    {
        using var ctx = DocumentContext<Presentation>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters = BuildParameters(operation, slideIndex, notes, slideIndices,
            headerText, footerText, dateText, showPageNumber);

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

        if (operation.ToLowerInvariant() == "get")
            return result;

        if (operationContext.IsModified)
            ctx.Save(outputPath);

        return $"{result}\n{ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Builds OperationParameters from method parameters.
    /// </summary>
    private static OperationParameters BuildParameters(
        string operation,
        int? slideIndex,
        string? notes,
        int[]? slideIndices,
        string? headerText,
        string? footerText,
        string? dateText,
        bool showPageNumber)
    {
        var parameters = new OperationParameters();

        switch (operation.ToLowerInvariant())
        {
            case "set":
                if (slideIndex.HasValue) parameters.Set("slideIndex", slideIndex.Value);
                if (notes != null) parameters.Set("notes", notes);
                break;

            case "get":
                if (slideIndex.HasValue) parameters.Set("slideIndex", slideIndex.Value);
                break;

            case "clear":
                if (slideIndices != null) parameters.Set("slideIndices", slideIndices);
                break;

            case "set_header_footer":
                if (headerText != null) parameters.Set("headerText", headerText);
                if (footerText != null) parameters.Set("footerText", footerText);
                if (dateText != null) parameters.Set("dateText", dateText);
                parameters.Set("showPageNumber", showPageNumber);
                break;
        }

        return parameters;
    }
}
