using System.ComponentModel;
using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Helpers;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Tool for managing PowerPoint handout settings (header/footer).
///     Note: Handout pages have separate header and footer fields (unlike slides which only have footer).
/// </summary>
[ToolHandlerMapping("AsposeMcpServer.Handlers.PowerPoint.Handout")]
[McpServerToolType]
public class PptHandoutTool
{
    /// <summary>
    ///     Handler registry for handout operations.
    /// </summary>
    private readonly HandlerRegistry<Presentation> _handlerRegistry;

    /// <summary>
    ///     Identity accessor for session isolation.
    /// </summary>
    private readonly ISessionIdentityAccessor? _identityAccessor;

    /// <summary>
    ///     Session manager for document lifecycle management.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="PptHandoutTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation.</param>
    public PptHandoutTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry =
            HandlerRegistry<Presentation>.CreateFromNamespace("AsposeMcpServer.Handlers.PowerPoint.Handout");
    }

    /// <summary>
    ///     Executes a PowerPoint handout operation (set_header_footer).
    /// </summary>
    /// <param name="operation">The operation to perform: set_header_footer.</param>
    /// <param name="path">Presentation file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="headerText">Header text for handout pages.</param>
    /// <param name="footerText">Footer text for handout pages.</param>
    /// <param name="dateText">Date/time text for handout pages.</param>
    /// <param name="showPageNumber">Show page number on handout pages (default: true).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when the operation is unknown.</exception>
    /// <exception cref="InvalidOperationException">Thrown when the presentation does not have a handout master slide.</exception>
    [McpServerTool(
        Name = "ppt_handout",
        Title = "PowerPoint Handout Operations",
        Destructive = false,
        Idempotent = true,
        OpenWorld = false,
        ReadOnly = false,
        UseStructuredContent = true)]
    [Description(@"Manage PowerPoint handout settings. Supports 1 operation: set_header_footer.

Note: Handout pages have separate header and footer fields (unlike slides which only have footer).
Important: Presentation must have a handout master (created via PowerPoint: View > Handout Master).

Usage examples:
- Set header/footer: ppt_handout(operation='set_header_footer', path='presentation.pptx', headerText='Header', footerText='Footer')")]
    public object Execute(
        [Description("Operation: set_header_footer")]
        string operation,
        [Description("Presentation file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Header text for handout pages")]
        string? headerText = null,
        [Description("Footer text for handout pages")]
        string? footerText = null,
        [Description("Date/time text for handout pages")]
        string? dateText = null,
        [Description("Show page number on handout pages (default: true)")]
        bool showPageNumber = true)
    {
        using var ctx = DocumentContext<Presentation>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters = BuildParameters(headerText, footerText, dateText, showPageNumber);

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

        if (operationContext.IsModified)
            ctx.Save(outputPath);

        return ResultHelper.FinalizeResult((dynamic)result, ctx, outputPath);
    }

    /// <summary>
    ///     Builds OperationParameters from method parameters.
    /// </summary>
    /// <param name="headerText">The header text for handout pages.</param>
    /// <param name="footerText">The footer text for handout pages.</param>
    /// <param name="dateText">The date/time text for handout pages.</param>
    /// <param name="showPageNumber">Whether to show page number on handout pages.</param>
    /// <returns>OperationParameters configured for the handout operation.</returns>
    private static OperationParameters BuildParameters(
        string? headerText,
        string? footerText,
        string? dateText,
        bool showPageNumber)
    {
        var parameters = new OperationParameters();

        if (headerText != null) parameters.Set("headerText", headerText);
        if (footerText != null) parameters.Set("footerText", footerText);
        if (dateText != null) parameters.Set("dateText", dateText);
        parameters.Set("showPageNumber", showPageNumber);

        return parameters;
    }
}
