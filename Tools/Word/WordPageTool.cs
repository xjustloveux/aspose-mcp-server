using System.ComponentModel;
using System.Text.Json.Nodes;
using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Helpers;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Unified tool for page operations in Word documents
///     Merges: WordSetPageMarginsTool, WordSetPageOrientationTool, WordSetPageSizeTool,
///     WordSetPageNumberTool, WordSetPageSetupTool, WordDeletePageTool, WordInsertBlankPageTool, WordAddPageBreakTool
/// </summary>
[ToolHandlerMapping("AsposeMcpServer.Handlers.Word.Page")]
[McpServerToolType]
public class WordPageTool
{
    /// <summary>
    ///     Handler registry for page operations
    /// </summary>
    private readonly HandlerRegistry<Document> _handlerRegistry;

    /// <summary>
    ///     Identity accessor for session isolation
    /// </summary>
    private readonly ISessionIdentityAccessor? _identityAccessor;

    /// <summary>
    ///     Session manager for document session operations
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the WordPageTool class
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document operations</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation</param>
    public WordPageTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry = HandlerRegistry<Document>.CreateFromNamespace("AsposeMcpServer.Handlers.Word.Page");
    }

    /// <summary>
    ///     Executes a Word page operation (set_margins, set_orientation, set_size, set_page_number, set_page_setup,
    ///     delete_page, insert_blank_page, add_page_break).
    /// </summary>
    /// <param name="operation">
    ///     The operation to perform: set_margins, set_orientation, set_size, set_page_number,
    ///     set_page_setup, delete_page, insert_blank_page, add_page_break.
    /// </param>
    /// <param name="path">Word document file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="top">Top margin in points (72 pts = 1 inch).</param>
    /// <param name="bottom">Bottom margin in points.</param>
    /// <param name="left">Left margin in points.</param>
    /// <param name="right">Right margin in points.</param>
    /// <param name="orientation">Page orientation: Portrait or Landscape.</param>
    /// <param name="width">Page width in points (72 pts = 1 inch).</param>
    /// <param name="height">Page height in points.</param>
    /// <param name="paperSize">Predefined paper size: A4, Letter, Legal, A3, A5.</param>
    /// <param name="pageNumberFormat">Page number format: arabic, roman, letter.</param>
    /// <param name="startingPageNumber">Starting page number.</param>
    /// <param name="sectionIndex">Section index (0-based).</param>
    /// <param name="sectionIndices">Array of section indices.</param>
    /// <param name="pageIndex">Page index to delete (0-based).</param>
    /// <param name="insertAtPageIndex">Page index to insert blank page at (0-based).</param>
    /// <param name="paragraphIndex">Paragraph index to insert page break after (0-based).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(
        Name = "word_page",
        Title = "Word Page Operations",
        Destructive = true,
        Idempotent = false,
        OpenWorld = false,
        ReadOnly = false,
        UseStructuredContent = true)]
    [Description(
        @"Manage page settings in Word documents. Supports 8 operations: set_margins, set_orientation, set_size, set_page_number, set_page_setup, delete_page, insert_blank_page, add_page_break.

Usage examples:
- Set margins: word_page(operation='set_margins', path='doc.docx', top=72, bottom=72, left=72, right=72)
- Set orientation: word_page(operation='set_orientation', path='doc.docx', orientation='landscape')
- Set page size: word_page(operation='set_size', path='doc.docx', width=792, height=612)
- Set page number: word_page(operation='set_page_number', path='doc.docx', startingPageNumber=1)
- Delete page: word_page(operation='delete_page', path='doc.docx', pageIndex=1)
- Insert blank page: word_page(operation='insert_blank_page', path='doc.docx', insertAtPageIndex=2)
- Add page break: word_page(operation='add_page_break', path='doc.docx', paragraphIndex=10)")]
    public object Execute(
        [Description(
            "Operation: set_margins, set_orientation, set_size, set_page_number, set_page_setup, delete_page, insert_blank_page, add_page_break")]
        string operation,
        [Description("Document file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Top margin in points (72 pts = 1 inch)")]
        double? top = null,
        [Description("Bottom margin in points")]
        double? bottom = null,
        [Description("Left margin in points")] double? left = null,
        [Description("Right margin in points")]
        double? right = null,
        [Description("Page orientation: Portrait or Landscape")]
        string? orientation = null,
        [Description("Page width in points (72 pts = 1 inch)")]
        double? width = null,
        [Description("Page height in points")] double? height = null,
        [Description("Predefined paper size: A4, Letter, Legal, A3, A5")]
        string? paperSize = null,
        [Description("Page number format: arabic, roman, letter")]
        string? pageNumberFormat = null,
        [Description("Starting page number")] int? startingPageNumber = null,
        [Description("Section index (0-based)")]
        int? sectionIndex = null,
        [Description("Array of section indices (overrides sectionIndex)")]
        JsonArray? sectionIndices = null,
        [Description("Page index to delete (0-based)")]
        int? pageIndex = null,
        [Description("Page index to insert blank page at (0-based)")]
        int? insertAtPageIndex = null,
        [Description("Paragraph index to insert page break after (0-based)")]
        int? paragraphIndex = null)
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters = BuildParameters(operation, top, bottom, left, right, orientation, width, height, paperSize,
            pageNumberFormat, startingPageNumber, sectionIndex, sectionIndices, pageIndex, insertAtPageIndex,
            paragraphIndex);

        var handler = _handlerRegistry.GetHandler(operation);

        var operationContext = new OperationContext<Document>
        {
            Document = ctx.Document,
            SessionManager = _sessionManager,
            IdentityAccessor = _identityAccessor,
            SessionId = sessionId,
            SourcePath = path,
            OutputPath = outputPath
        };

        var result = handler.Execute(operationContext, parameters);

        // Handle delete_page operation which creates a new document
        if (string.Equals(operation, "delete_page", StringComparison.OrdinalIgnoreCase) &&
            operationContext.ResultDocument != null)
        {
            if (!ctx.IsSession)
            {
                var savePath = outputPath ?? throw new InvalidOperationException("Output path required for file mode");
                operationContext.ResultDocument.Save(savePath);
                return ResultHelper.FinalizeResult((dynamic)$"{result}\nOutput: {savePath}", savePath, sessionId);
            }

            // For session mode, update the session document
            var doc = ctx.Document;
            doc.RemoveAllChildren();
            foreach (var section in operationContext.ResultDocument.Sections.Cast<Section>())
                doc.AppendChild(doc.ImportNode(section, true));
        }

        // Standard save behavior
        if (operationContext.IsModified)
            ctx.Save(outputPath);

        return ResultHelper.FinalizeResult((dynamic)result, ctx, outputPath);
    }

    /// <summary>
    ///     Builds OperationParameters from method parameters using strategy pattern.
    ///     Parameters are documented on the Execute method.
    /// </summary>
    /// <returns>OperationParameters configured with all input values.</returns>
    private static OperationParameters BuildParameters(
        string operation,
        double? top,
        double? bottom,
        double? left,
        double? right,
        string? orientation,
        double? width,
        double? height,
        string? paperSize,
        string? pageNumberFormat,
        int? startingPageNumber,
        int? sectionIndex,
        JsonArray? sectionIndices,
        int? pageIndex,
        int? insertAtPageIndex,
        int? paragraphIndex)
    {
        var parameters = new OperationParameters();

        return operation.ToLower() switch
        {
            "set_margins" => BuildSetMarginsParameters(parameters, top, bottom, left, right, sectionIndex,
                sectionIndices),
            "set_orientation" => BuildSetOrientationParameters(parameters, orientation, sectionIndex, sectionIndices),
            "set_size" => BuildSetSizeParameters(parameters, width, height, paperSize, sectionIndex, sectionIndices),
            "set_page_number" => BuildSetPageNumberParameters(parameters, pageNumberFormat, startingPageNumber,
                sectionIndex),
            "set_page_setup" => BuildSetPageSetupParameters(parameters, top, bottom, left, right, orientation,
                sectionIndex),
            "delete_page" => BuildDeletePageParameters(parameters, pageIndex),
            "insert_blank_page" => BuildInsertBlankPageParameters(parameters, insertAtPageIndex),
            "add_page_break" => BuildAddPageBreakParameters(parameters, paragraphIndex),
            _ => parameters
        };
    }

    /// <summary>
    ///     Builds parameters for the set margins operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="top">The top margin in points (72 pts = 1 inch).</param>
    /// <param name="bottom">The bottom margin in points.</param>
    /// <param name="left">The left margin in points.</param>
    /// <param name="right">The right margin in points.</param>
    /// <param name="sectionIndex">The section index (0-based).</param>
    /// <param name="sectionIndices">The array of section indices.</param>
    /// <returns>OperationParameters configured for the set margins operation.</returns>
    private static OperationParameters BuildSetMarginsParameters(OperationParameters parameters, double? top,
        double? bottom, double? left, double? right, int? sectionIndex, JsonArray? sectionIndices)
    {
        if (top.HasValue) parameters.Set("top", top.Value);
        if (bottom.HasValue) parameters.Set("bottom", bottom.Value);
        if (left.HasValue) parameters.Set("left", left.Value);
        if (right.HasValue) parameters.Set("right", right.Value);
        if (sectionIndex.HasValue) parameters.Set("sectionIndex", sectionIndex.Value);
        if (sectionIndices != null) parameters.Set("sectionIndices", sectionIndices);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the set orientation operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="orientation">The page orientation: 'Portrait' or 'Landscape'.</param>
    /// <param name="sectionIndex">The section index (0-based).</param>
    /// <param name="sectionIndices">The array of section indices.</param>
    /// <returns>OperationParameters configured for the set orientation operation.</returns>
    private static OperationParameters BuildSetOrientationParameters(OperationParameters parameters,
        string? orientation, int? sectionIndex, JsonArray? sectionIndices)
    {
        if (orientation != null) parameters.Set("orientation", orientation);
        if (sectionIndex.HasValue) parameters.Set("sectionIndex", sectionIndex.Value);
        if (sectionIndices != null) parameters.Set("sectionIndices", sectionIndices);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the set page size operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="width">The page width in points.</param>
    /// <param name="height">The page height in points.</param>
    /// <param name="paperSize">The predefined paper size: 'A4', 'Letter', 'Legal', 'A3', 'A5'.</param>
    /// <param name="sectionIndex">The section index (0-based).</param>
    /// <param name="sectionIndices">The array of section indices.</param>
    /// <returns>OperationParameters configured for the set size operation.</returns>
    private static OperationParameters BuildSetSizeParameters(OperationParameters parameters, double? width,
        double? height, string? paperSize, int? sectionIndex, JsonArray? sectionIndices)
    {
        if (width.HasValue) parameters.Set("width", width.Value);
        if (height.HasValue) parameters.Set("height", height.Value);
        if (paperSize != null) parameters.Set("paperSize", paperSize);
        if (sectionIndex.HasValue) parameters.Set("sectionIndex", sectionIndex.Value);
        if (sectionIndices != null) parameters.Set("sectionIndices", sectionIndices);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the set page number operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="pageNumberFormat">The page number format: 'arabic', 'roman', 'letter'.</param>
    /// <param name="startingPageNumber">The starting page number.</param>
    /// <param name="sectionIndex">The section index (0-based).</param>
    /// <returns>OperationParameters configured for the set page number operation.</returns>
    private static OperationParameters BuildSetPageNumberParameters(OperationParameters parameters,
        string? pageNumberFormat, int? startingPageNumber, int? sectionIndex)
    {
        if (pageNumberFormat != null) parameters.Set("pageNumberFormat", pageNumberFormat);
        if (startingPageNumber.HasValue) parameters.Set("startingPageNumber", startingPageNumber.Value);
        if (sectionIndex.HasValue) parameters.Set("sectionIndex", sectionIndex.Value);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the set page setup operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="top">The top margin in points.</param>
    /// <param name="bottom">The bottom margin in points.</param>
    /// <param name="left">The left margin in points.</param>
    /// <param name="right">The right margin in points.</param>
    /// <param name="orientation">The page orientation: 'Portrait' or 'Landscape'.</param>
    /// <param name="sectionIndex">The section index (0-based).</param>
    /// <returns>OperationParameters configured for the set page setup operation.</returns>
    private static OperationParameters BuildSetPageSetupParameters(OperationParameters parameters, double? top,
        double? bottom, double? left, double? right, string? orientation, int? sectionIndex)
    {
        if (top.HasValue) parameters.Set("top", top.Value);
        if (bottom.HasValue) parameters.Set("bottom", bottom.Value);
        if (left.HasValue) parameters.Set("left", left.Value);
        if (right.HasValue) parameters.Set("right", right.Value);
        if (orientation != null) parameters.Set("orientation", orientation);
        if (sectionIndex.HasValue) parameters.Set("sectionIndex", sectionIndex.Value);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the delete page operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="pageIndex">The page index to delete (0-based).</param>
    /// <returns>OperationParameters configured for the delete page operation.</returns>
    private static OperationParameters BuildDeletePageParameters(OperationParameters parameters, int? pageIndex)
    {
        if (pageIndex.HasValue) parameters.Set("pageIndex", pageIndex.Value);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the insert blank page operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="insertAtPageIndex">The page index to insert blank page at (0-based).</param>
    /// <returns>OperationParameters configured for the insert blank page operation.</returns>
    private static OperationParameters BuildInsertBlankPageParameters(OperationParameters parameters,
        int? insertAtPageIndex)
    {
        if (insertAtPageIndex.HasValue) parameters.Set("insertAtPageIndex", insertAtPageIndex.Value);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the add page break operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="paragraphIndex">The paragraph index to insert page break after (0-based).</param>
    /// <returns>OperationParameters configured for the add page break operation.</returns>
    private static OperationParameters BuildAddPageBreakParameters(OperationParameters parameters, int? paragraphIndex)
    {
        if (paragraphIndex.HasValue) parameters.Set("paragraphIndex", paragraphIndex.Value);
        return parameters;
    }
}
