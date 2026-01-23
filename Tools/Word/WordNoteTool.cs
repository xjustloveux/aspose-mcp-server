using System.ComponentModel;
using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Helpers;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Unified tool for footnote and endnote operations in Word documents
///     Merges: WordAddFootnoteTool, WordAddEndnoteTool, WordDeleteFootnoteTool, WordDeleteEndnoteTool,
///     WordEditFootnoteTool, WordEditEndnoteTool, WordGetFootnotesTool, WordGetEndnotesTool
/// </summary>
[ToolHandlerMapping("AsposeMcpServer.Handlers.Word.Note")]
[McpServerToolType]
public class WordNoteTool
{
    /// <summary>
    ///     Handler registry for note operations
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
    ///     Initializes a new instance of the WordNoteTool class
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document operations</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation</param>
    public WordNoteTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry = HandlerRegistry<Document>.CreateFromNamespace("AsposeMcpServer.Handlers.Word.Note");
    }

    /// <summary>
    ///     Executes a Word note operation (add_footnote, add_endnote, delete_footnote, delete_endnote, edit_footnote,
    ///     edit_endnote, get_footnotes, get_endnotes).
    /// </summary>
    /// <param name="operation">
    ///     The operation to perform: add_footnote, add_endnote, delete_footnote, delete_endnote,
    ///     edit_footnote, edit_endnote, get_footnotes, get_endnotes.
    /// </param>
    /// <param name="path">Word document file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="text">Note text content.</param>
    /// <param name="paragraphIndex">Paragraph index (0-based, -1 for document end).</param>
    /// <param name="sectionIndex">Section index (0-based, default: 0).</param>
    /// <param name="referenceText">Reference text in document to insert note at.</param>
    /// <param name="customMark">Custom note mark.</param>
    /// <param name="referenceMark">Reference mark of note to delete/edit.</param>
    /// <param name="noteIndex">Note index (0-based).</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(
        Name = "word_note",
        Title = "Word Note Operations",
        Destructive = true,
        Idempotent = false,
        OpenWorld = false,
        ReadOnly = false,
        UseStructuredContent = true)]
    [Description(
        @"Manage footnotes and endnotes in Word documents. Supports 8 operations: add_footnote, add_endnote, delete_footnote, delete_endnote, edit_footnote, edit_endnote, get_footnotes, get_endnotes.

Usage examples:
- Add footnote: word_note(operation='add_footnote', path='doc.docx', text='Footnote text', paragraphIndex=0)
- Add endnote: word_note(operation='add_endnote', path='doc.docx', text='Endnote text', paragraphIndex=0)
- Delete footnote: word_note(operation='delete_footnote', path='doc.docx', noteIndex=0)
- Edit footnote: word_note(operation='edit_footnote', path='doc.docx', noteIndex=0, text='Updated footnote')
- Get footnotes: word_note(operation='get_footnotes', path='doc.docx')")]
    public object Execute(
        [Description(
            "Operation: add_footnote, add_endnote, delete_footnote, delete_endnote, edit_footnote, edit_endnote, get_footnotes, get_endnotes")]
        string operation,
        [Description("Document file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Note text content")] string? text = null,
        [Description("Paragraph index (0-based, -1 for document end)")]
        int? paragraphIndex = null,
        [Description("Section index (0-based, default: 0)")]
        int sectionIndex = 0,
        [Description("Reference text in document to insert note at")]
        string? referenceText = null,
        [Description("Custom note mark")] string? customMark = null,
        [Description("Reference mark of note to delete/edit")]
        string? referenceMark = null,
        [Description("Note index (0-based)")] int? noteIndex = null)
    {
        var parameters = BuildParameters(operation, text, paragraphIndex, sectionIndex, referenceText, customMark,
            referenceMark, noteIndex);

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

        return ResultHelper.FinalizeResult((dynamic)result, ctx, effectiveOutputPath);
    }

    /// <summary>
    ///     Builds OperationParameters from method parameters using strategy pattern.
    ///     Parameters are documented on the Execute method.
    /// </summary>
    /// <returns>OperationParameters configured with all input values.</returns>
    private static OperationParameters BuildParameters(
        string operation,
        string? text,
        int? paragraphIndex,
        int sectionIndex,
        string? referenceText,
        string? customMark,
        string? referenceMark,
        int? noteIndex)
    {
        var parameters = new OperationParameters();

        return operation.ToLower() switch
        {
            "add_footnote" or "add_endnote" => BuildAddNoteParameters(parameters, text, paragraphIndex, sectionIndex,
                referenceText, customMark),
            "delete_footnote" or "delete_endnote" => BuildDeleteNoteParameters(parameters, referenceMark, noteIndex),
            "edit_footnote" or "edit_endnote" => BuildEditNoteParameters(parameters, referenceMark, noteIndex, text),
            _ => parameters
        };
    }

    /// <summary>
    ///     Builds parameters for the add footnote/endnote operations.
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="text">The note text content.</param>
    /// <param name="paragraphIndex">The paragraph index (0-based, -1 for document end).</param>
    /// <param name="sectionIndex">The section index (0-based).</param>
    /// <param name="referenceText">The reference text in document to insert note at.</param>
    /// <param name="customMark">The custom note mark.</param>
    /// <returns>OperationParameters configured for the add note operation.</returns>
    private static OperationParameters BuildAddNoteParameters(OperationParameters parameters, string? text,
        int? paragraphIndex, int sectionIndex, string? referenceText, string? customMark)
    {
        if (text != null) parameters.Set("text", text);
        if (paragraphIndex.HasValue) parameters.Set("paragraphIndex", paragraphIndex.Value);
        parameters.Set("sectionIndex", sectionIndex);
        if (referenceText != null) parameters.Set("referenceText", referenceText);
        if (customMark != null) parameters.Set("customMark", customMark);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the delete footnote/endnote operations.
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="referenceMark">The reference mark of the note to delete.</param>
    /// <param name="noteIndex">The note index (0-based).</param>
    /// <returns>OperationParameters configured for the delete note operation.</returns>
    private static OperationParameters BuildDeleteNoteParameters(OperationParameters parameters, string? referenceMark,
        int? noteIndex)
    {
        if (referenceMark != null) parameters.Set("referenceMark", referenceMark);
        if (noteIndex.HasValue) parameters.Set("noteIndex", noteIndex.Value);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the edit footnote/endnote operations.
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="referenceMark">The reference mark of the note to edit.</param>
    /// <param name="noteIndex">The note index (0-based).</param>
    /// <param name="text">The updated note text content.</param>
    /// <returns>OperationParameters configured for the edit note operation.</returns>
    private static OperationParameters BuildEditNoteParameters(OperationParameters parameters, string? referenceMark,
        int? noteIndex, string? text)
    {
        if (referenceMark != null) parameters.Set("referenceMark", referenceMark);
        if (noteIndex.HasValue) parameters.Set("noteIndex", noteIndex.Value);
        if (text != null) parameters.Set("text", text);
        return parameters;
    }
}
