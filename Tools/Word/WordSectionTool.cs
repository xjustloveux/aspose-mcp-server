using System.ComponentModel;
using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Unified tool for managing Word sections (insert, delete, get info)
///     Merges: WordInsertSectionTool, WordDeleteSectionTool, WordGetSectionsTool, WordGetSectionsInfoTool
/// </summary>
[McpServerToolType]
public class WordSectionTool
{
    /// <summary>
    ///     Handler registry for section operations
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
    ///     Initializes a new instance of the WordSectionTool class
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document operations</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation</param>
    public WordSectionTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry = HandlerRegistry<Document>.CreateFromNamespace("AsposeMcpServer.Handlers.Word.SectionBreak");
    }

    /// <summary>
    ///     Executes a Word section operation (insert, delete, get).
    /// </summary>
    /// <param name="operation">The operation to perform: insert, delete, get.</param>
    /// <param name="path">Word document file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="sectionBreakType">Section break type: NextPage, Continuous, EvenPage, OddPage (for insert).</param>
    /// <param name="insertAtParagraphIndex">Paragraph index to insert section break after (0-based, for insert).</param>
    /// <param name="sectionIndex">Section index (0-based, for insert/delete/get).</param>
    /// <param name="sectionIndices">Array of section indices to delete (0-based, for delete).</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "word_section")]
    [Description(@"Manage Word document sections. Supports 3 operations: insert, delete, get.

Usage examples:
- Insert section: word_section(operation='insert', path='doc.docx', sectionBreakType='NextPage', insertAtParagraphIndex=5)
- Delete section: word_section(operation='delete', path='doc.docx', sectionIndex=1)
- Get sections: word_section(operation='get', path='doc.docx')

Notes:
- Section break types: NextPage (new page), Continuous (same page), EvenPage, OddPage
- IMPORTANT: Deleting a section will also delete all content within that section (paragraphs, tables, images)
- Use 'get' operation first to see section indices and their content statistics before deleting")]
    public string Execute( // NOSONAR S107 - MCP protocol requires multiple parameters
        [Description("Operation: insert, delete, get")]
        string operation,
        [Description("Document file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Section break type: NextPage, Continuous, EvenPage, OddPage (for insert)")]
        string? sectionBreakType = null,
        [Description("Paragraph index to insert section break after (0-based, for insert)")]
        int? insertAtParagraphIndex = null,
        [Description("Section index (0-based, for insert/delete/get)")]
        int? sectionIndex = null,
        [Description("Array of section indices to delete (0-based, overrides sectionIndex, for delete)")]
        int[]? sectionIndices = null)
    {
        var parameters = BuildParameters(operation, sectionBreakType, insertAtParagraphIndex, sectionIndex,
            sectionIndices);

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
        string? sectionBreakType,
        int? insertAtParagraphIndex,
        int? sectionIndex,
        int[]? sectionIndices)
    {
        return operation.ToLower() switch
        {
            "insert" => BuildInsertParameters(sectionBreakType, insertAtParagraphIndex, sectionIndex),
            "delete" => BuildDeleteParameters(sectionIndex, sectionIndices),
            "get" => BuildGetParameters(sectionIndex),
            _ => new OperationParameters()
        };
    }

    /// <summary>
    ///     Builds parameters for the insert section operation.
    /// </summary>
    /// <param name="sectionBreakType">The section break type (NextPage, Continuous, etc.).</param>
    /// <param name="insertAtParagraphIndex">The paragraph index to insert section break after.</param>
    /// <param name="sectionIndex">The section index.</param>
    /// <returns>OperationParameters configured for inserting a section.</returns>
    private static OperationParameters BuildInsertParameters(string? sectionBreakType, int? insertAtParagraphIndex,
        int? sectionIndex)
    {
        var parameters = new OperationParameters();
        if (sectionBreakType != null) parameters.Set("sectionBreakType", sectionBreakType);
        if (insertAtParagraphIndex.HasValue) parameters.Set("insertAtParagraphIndex", insertAtParagraphIndex.Value);
        if (sectionIndex.HasValue) parameters.Set("sectionIndex", sectionIndex.Value);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the delete section operation.
    /// </summary>
    /// <param name="sectionIndex">The section index to delete (0-based).</param>
    /// <param name="sectionIndices">Array of section indices to delete.</param>
    /// <returns>OperationParameters configured for deleting sections.</returns>
    private static OperationParameters BuildDeleteParameters(int? sectionIndex, int[]? sectionIndices)
    {
        var parameters = new OperationParameters();
        if (sectionIndex.HasValue) parameters.Set("sectionIndex", sectionIndex.Value);
        if (sectionIndices != null) parameters.Set("sectionIndices", sectionIndices);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the get section info operation.
    /// </summary>
    /// <param name="sectionIndex">The section index to get info for (0-based, optional).</param>
    /// <returns>OperationParameters configured for getting section info.</returns>
    private static OperationParameters BuildGetParameters(int? sectionIndex)
    {
        var parameters = new OperationParameters();
        if (sectionIndex.HasValue) parameters.Set("sectionIndex", sectionIndex.Value);
        return parameters;
    }
}
