using System.ComponentModel;
using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Tool for managing cross-references in Word documents
/// </summary>
[McpServerToolType]
public class WordReferenceTool
{
    /// <summary>
    ///     Handler registry for reference operations
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
    ///     Initializes a new instance of the WordReferenceTool class
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document operations</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation</param>
    public WordReferenceTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry = HandlerRegistry<Document>.CreateFromNamespace("AsposeMcpServer.Handlers.Word.Reference");
    }

    /// <summary>
    ///     Executes a Word reference operation (add_table_of_contents, update_table_of_contents, add_index,
    ///     add_cross_reference).
    /// </summary>
    /// <param name="operation">
    ///     The operation to perform: add_table_of_contents, update_table_of_contents, add_index,
    ///     add_cross_reference.
    /// </param>
    /// <param name="path">Word document file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (optional, defaults to overwrite input).</param>
    /// <param name="position">Insert position: start, end (for add_table_of_contents, default: start).</param>
    /// <param name="title">Table of contents title (for add_table_of_contents).</param>
    /// <param name="maxLevel">Maximum heading level to include (for add_table_of_contents, default: 3).</param>
    /// <param name="hyperlinks">Enable clickable hyperlinks (for add_table_of_contents, default: true).</param>
    /// <param name="pageNumbers">Show page numbers (for add_table_of_contents, default: true).</param>
    /// <param name="rightAlignPageNumbers">Right-align page numbers (for add_table_of_contents, default: true).</param>
    /// <param name="tocIndex">TOC field index (0-based, for update_table_of_contents).</param>
    /// <param name="indexEntries">Array of index entries as JSON string (for add_index).</param>
    /// <param name="insertIndexAtEnd">Insert INDEX field at end of document (for add_index, default: true).</param>
    /// <param name="headingStyle">Heading style for index (for add_index, default: 'Heading 1').</param>
    /// <param name="referenceType">Reference type: Heading, Bookmark, Figure, Table, Equation (for add_cross_reference).</param>
    /// <param name="referenceText">Text to insert before reference (for add_cross_reference).</param>
    /// <param name="targetName">Target name (heading text, bookmark name, etc.) (for add_cross_reference).</param>
    /// <param name="insertAsHyperlink">Insert as hyperlink (for add_cross_reference, default: true).</param>
    /// <param name="includeAboveBelow">Include 'above' or 'below' text (for add_cross_reference, default: false).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "word_reference")]
    [Description(
        @"Manage references in Word documents. Supports 4 operations: add_table_of_contents, update_table_of_contents, add_index, add_cross_reference.

Usage examples:
- Add table of contents: word_reference(operation='add_table_of_contents', path='doc.docx', title='Table of Contents', maxLevel=3)
- Update table of contents: word_reference(operation='update_table_of_contents', path='doc.docx')
- Add index: word_reference(operation='add_index', path='doc.docx', indexEntries='[{""text"":""Index term""}]')
- Add cross-reference: word_reference(operation='add_cross_reference', path='doc.docx', referenceType='Bookmark', targetName='Chapter1', referenceText='See ')

Notes:
- TOC is automatically updated after insertion using UpdateFields()
- For cross-references, targetName must be an existing bookmark name in the document
- If headingStyle doesn't exist in the document, it falls back to 'Heading 1'")]
    public string Execute( // NOSONAR S107 - MCP protocol requires multiple parameters
        [Description("Operation: add_table_of_contents, update_table_of_contents, add_index, add_cross_reference")]
        string operation,
        [Description("Document file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (optional, defaults to overwrite input)")]
        string? outputPath = null,
        [Description("Insert position: start, end (for add_table_of_contents, default: start)")]
        string position = "start",
        [Description("Table of contents title (for add_table_of_contents, default: 'Table of Contents')")]
        string title = "Table of Contents",
        [Description("Maximum heading level to include (for add_table_of_contents, default: 3)")]
        int maxLevel = 3,
        [Description("Enable clickable hyperlinks (for add_table_of_contents, default: true)")]
        bool hyperlinks = true,
        [Description("Show page numbers (for add_table_of_contents, default: true)")]
        bool pageNumbers = true,
        [Description("Right-align page numbers (for add_table_of_contents, default: true)")]
        bool rightAlignPageNumbers = true,
        [Description("TOC field index (0-based, for update_table_of_contents, optional)")]
        int? tocIndex = null,
        [Description("Array of index entries as JSON string (for add_index)")]
        string? indexEntries = null,
        [Description("Insert INDEX field at end of document (for add_index, default: true)")]
        bool insertIndexAtEnd = true,
        [Description("Heading style for index (for add_index, default: 'Heading 1')")]
        string headingStyle = "Heading 1",
        [Description("Reference type: Heading, Bookmark, Figure, Table, Equation (for add_cross_reference)")]
        string? referenceType = null,
        [Description("Text to insert before reference (for add_cross_reference, optional)")]
        string? referenceText = null,
        [Description("Target name (heading text, bookmark name, etc.) (for add_cross_reference)")]
        string? targetName = null,
        [Description("Insert as hyperlink (for add_cross_reference, default: true)")]
        bool insertAsHyperlink = true,
        [Description("Include 'above' or 'below' text (for add_cross_reference, default: false)")]
        bool includeAboveBelow = false)
    {
        var parameters = BuildParameters(operation, position, title, maxLevel, hyperlinks, pageNumbers,
            rightAlignPageNumbers, tocIndex, indexEntries, insertIndexAtEnd, headingStyle, referenceType,
            referenceText, targetName, insertAsHyperlink, includeAboveBelow);

        var handler = _handlerRegistry.GetHandler(operation);

        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var effectiveOutputPath = outputPath ?? path;

        var operationContext = new OperationContext<Document>
        {
            Document = ctx.Document,
            SessionManager = _sessionManager,
            IdentityAccessor = _identityAccessor,
            SessionId = sessionId,
            SourcePath = path,
            OutputPath = effectiveOutputPath
        };

        var result = handler.Execute(operationContext, parameters);

        if (operationContext.IsModified)
            ctx.Save(effectiveOutputPath);

        return ctx.IsSession ? result : $"{result}\n{ctx.GetOutputMessage(effectiveOutputPath)}";
    }

    /// <summary>
    ///     Builds OperationParameters from method parameters using strategy pattern.
    ///     Parameters are documented on the Execute method.
    /// </summary>
    /// <returns>OperationParameters configured with all input values.</returns>
    private static OperationParameters BuildParameters( // NOSONAR S107 - MCP protocol parameter building
        string operation,
        string position,
        string title,
        int maxLevel,
        bool hyperlinks,
        bool pageNumbers,
        bool rightAlignPageNumbers,
        int? tocIndex,
        string? indexEntries,
        bool insertIndexAtEnd,
        string headingStyle,
        string? referenceType,
        string? referenceText,
        string? targetName,
        bool insertAsHyperlink,
        bool includeAboveBelow)
    {
        return operation.ToLower() switch
        {
            "add_table_of_contents" => BuildAddTableOfContentsParameters(position, title, maxLevel, hyperlinks,
                pageNumbers, rightAlignPageNumbers),
            "update_table_of_contents" => BuildUpdateTableOfContentsParameters(tocIndex),
            "add_index" => BuildAddIndexParameters(indexEntries, insertIndexAtEnd, headingStyle),
            "add_cross_reference" => BuildAddCrossReferenceParameters(referenceType, referenceText, targetName,
                insertAsHyperlink, includeAboveBelow),
            _ => new OperationParameters()
        };
    }

    /// <summary>
    ///     Builds parameters for the add_table_of_contents operation.
    /// </summary>
    /// <param name="position">The insert position: start, end.</param>
    /// <param name="title">The table of contents title.</param>
    /// <param name="maxLevel">The maximum heading level to include.</param>
    /// <param name="hyperlinks">Whether to enable clickable hyperlinks.</param>
    /// <param name="pageNumbers">Whether to show page numbers.</param>
    /// <param name="rightAlignPageNumbers">Whether to right-align page numbers.</param>
    /// <returns>OperationParameters configured for adding table of contents.</returns>
    private static OperationParameters BuildAddTableOfContentsParameters(string position, string title, int maxLevel,
        bool hyperlinks, bool pageNumbers, bool rightAlignPageNumbers)
    {
        var parameters = new OperationParameters();
        parameters.Set("position", position);
        parameters.Set("title", title);
        parameters.Set("maxLevel", maxLevel);
        parameters.Set("hyperlinks", hyperlinks);
        parameters.Set("pageNumbers", pageNumbers);
        parameters.Set("rightAlignPageNumbers", rightAlignPageNumbers);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the update_table_of_contents operation.
    /// </summary>
    /// <param name="tocIndex">The TOC field index (0-based).</param>
    /// <returns>OperationParameters configured for updating table of contents.</returns>
    private static OperationParameters BuildUpdateTableOfContentsParameters(int? tocIndex)
    {
        var parameters = new OperationParameters();
        if (tocIndex.HasValue) parameters.Set("tocIndex", tocIndex.Value);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the add_index operation.
    /// </summary>
    /// <param name="indexEntries">The array of index entries as JSON string.</param>
    /// <param name="insertIndexAtEnd">Whether to insert INDEX field at end of document.</param>
    /// <param name="headingStyle">The heading style for index.</param>
    /// <returns>OperationParameters configured for adding index.</returns>
    private static OperationParameters BuildAddIndexParameters(string? indexEntries, bool insertIndexAtEnd,
        string headingStyle)
    {
        var parameters = new OperationParameters();
        if (indexEntries != null) parameters.Set("indexEntries", indexEntries);
        parameters.Set("insertIndexAtEnd", insertIndexAtEnd);
        parameters.Set("headingStyle", headingStyle);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the add_cross_reference operation.
    /// </summary>
    /// <param name="referenceType">The reference type: Heading, Bookmark, Figure, Table, Equation.</param>
    /// <param name="referenceText">The text to insert before reference.</param>
    /// <param name="targetName">The target name (heading text, bookmark name, etc.).</param>
    /// <param name="insertAsHyperlink">Whether to insert as hyperlink.</param>
    /// <param name="includeAboveBelow">Whether to include 'above' or 'below' text.</param>
    /// <returns>OperationParameters configured for adding cross-reference.</returns>
    private static OperationParameters BuildAddCrossReferenceParameters(string? referenceType, string? referenceText,
        string? targetName, bool insertAsHyperlink, bool includeAboveBelow)
    {
        var parameters = new OperationParameters();
        if (referenceType != null) parameters.Set("referenceType", referenceType);
        if (referenceText != null) parameters.Set("referenceText", referenceText);
        if (targetName != null) parameters.Set("targetName", targetName);
        parameters.Set("insertAsHyperlink", insertAsHyperlink);
        parameters.Set("includeAboveBelow", includeAboveBelow);
        return parameters;
    }
}
