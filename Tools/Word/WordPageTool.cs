using System.ComponentModel;
using System.Text.Json.Nodes;
using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Unified tool for page operations in Word documents
///     Merges: WordSetPageMarginsTool, WordSetPageOrientationTool, WordSetPageSizeTool,
///     WordSetPageNumberTool, WordSetPageSetupTool, WordDeletePageTool, WordInsertBlankPageTool, WordAddPageBreakTool
/// </summary>
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
    [McpServerTool(Name = "word_page")]
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
    public string Execute(
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
                return $"{result}\nOutput: {savePath}";
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

        return $"{result}\n{ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Builds OperationParameters from method parameters.
    /// </summary>
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

        switch (operation.ToLower())
        {
            case "set_margins":
                if (top.HasValue) parameters.Set("top", top.Value);
                if (bottom.HasValue) parameters.Set("bottom", bottom.Value);
                if (left.HasValue) parameters.Set("left", left.Value);
                if (right.HasValue) parameters.Set("right", right.Value);
                if (sectionIndex.HasValue) parameters.Set("sectionIndex", sectionIndex.Value);
                if (sectionIndices != null) parameters.Set("sectionIndices", sectionIndices);
                break;

            case "set_orientation":
                if (orientation != null) parameters.Set("orientation", orientation);
                if (sectionIndex.HasValue) parameters.Set("sectionIndex", sectionIndex.Value);
                if (sectionIndices != null) parameters.Set("sectionIndices", sectionIndices);
                break;

            case "set_size":
                if (width.HasValue) parameters.Set("width", width.Value);
                if (height.HasValue) parameters.Set("height", height.Value);
                if (paperSize != null) parameters.Set("paperSize", paperSize);
                if (sectionIndex.HasValue) parameters.Set("sectionIndex", sectionIndex.Value);
                if (sectionIndices != null) parameters.Set("sectionIndices", sectionIndices);
                break;

            case "set_page_number":
                if (pageNumberFormat != null) parameters.Set("pageNumberFormat", pageNumberFormat);
                if (startingPageNumber.HasValue) parameters.Set("startingPageNumber", startingPageNumber.Value);
                if (sectionIndex.HasValue) parameters.Set("sectionIndex", sectionIndex.Value);
                break;

            case "set_page_setup":
                if (top.HasValue) parameters.Set("top", top.Value);
                if (bottom.HasValue) parameters.Set("bottom", bottom.Value);
                if (left.HasValue) parameters.Set("left", left.Value);
                if (right.HasValue) parameters.Set("right", right.Value);
                if (orientation != null) parameters.Set("orientation", orientation);
                if (sectionIndex.HasValue) parameters.Set("sectionIndex", sectionIndex.Value);
                break;

            case "delete_page":
                if (pageIndex.HasValue) parameters.Set("pageIndex", pageIndex.Value);
                break;

            case "insert_blank_page":
                if (insertAtPageIndex.HasValue) parameters.Set("insertAtPageIndex", insertAtPageIndex.Value);
                break;

            case "add_page_break":
                if (paragraphIndex.HasValue) parameters.Set("paragraphIndex", paragraphIndex.Value);
                break;
        }

        return parameters;
    }
}
