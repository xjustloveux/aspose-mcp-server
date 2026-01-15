using System.ComponentModel;
using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Unified tool for text operations in Word documents.
///     Dispatches operations to individual handlers via HandlerRegistry.
/// </summary>
[McpServerToolType]
public class WordTextTool
{
    /// <summary>
    ///     Registry of text operation handlers.
    /// </summary>
    private readonly HandlerRegistry<Document> _handlerRegistry;

    /// <summary>
    ///     Identity accessor for session isolation.
    /// </summary>
    private readonly ISessionIdentityAccessor? _identityAccessor;

    /// <summary>
    ///     Session manager for document session operations.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the WordTextTool class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document operations.</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation.</param>
    public WordTextTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry = HandlerRegistry<Document>.CreateFromNamespace("AsposeMcpServer.Handlers.Word.Text");
    }

    /// <summary>
    ///     Executes a Word text operation (add, delete, replace, search, format, insert_at_position, delete_range,
    ///     add_with_style).
    /// </summary>
    /// <param name="operation">
    ///     The operation to perform: add, delete, replace, search, format, insert_at_position,
    ///     delete_range, add_with_style.
    /// </param>
    /// <param name="path">Word document file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (for write operations).</param>
    /// <param name="text">Text content (for add, replace, insert_at_position, add_with_style).</param>
    /// <param name="fontName">Font name.</param>
    /// <param name="fontNameAscii">Font name for ASCII characters.</param>
    /// <param name="fontNameFarEast">Font name for Far East characters.</param>
    /// <param name="fontSize">Font size.</param>
    /// <param name="bold">Bold text.</param>
    /// <param name="italic">Italic text.</param>
    /// <param name="underline">Underline style: none, single, double, dotted, dash.</param>
    /// <param name="color">Text color hex or name.</param>
    /// <param name="strikethrough">Strikethrough text.</param>
    /// <param name="superscript">Superscript text.</param>
    /// <param name="subscript">Subscript text.</param>
    /// <param name="find">Text to find (for replace).</param>
    /// <param name="replace">Replacement text (for replace).</param>
    /// <param name="searchText">Text to search for (for search/delete).</param>
    /// <param name="caseSensitive">Case sensitive search (for search/replace).</param>
    /// <param name="useRegex">Use regular expression (for search/replace).</param>
    /// <param name="replaceInFields">Replace text inside fields (for replace).</param>
    /// <param name="maxResults">Maximum number of results to return (for search).</param>
    /// <param name="contextLength">Number of characters to show for context (for search).</param>
    /// <param name="paragraphIndex">Paragraph index (0-based).</param>
    /// <param name="runIndex">Run index within paragraph (0-based).</param>
    /// <param name="startParagraphIndex">Start paragraph index (for delete_range).</param>
    /// <param name="startRunIndex">Start run index (for delete_range).</param>
    /// <param name="endParagraphIndex">End paragraph index (for delete_range).</param>
    /// <param name="endRunIndex">End run index (for delete_range).</param>
    /// <param name="insertParagraphIndex">Paragraph index for insert_at_position.</param>
    /// <param name="charIndex">Character index for insert_at_position.</param>
    /// <param name="sectionIndex">Section index (0-based).</param>
    /// <param name="insertBefore">Insert before position (for insert_at_position).</param>
    /// <param name="startCharIndex">Start character index (for delete_range).</param>
    /// <param name="endCharIndex">End character index (for delete_range).</param>
    /// <param name="styleName">Style name (for add_with_style).</param>
    /// <param name="alignment">Text alignment (for add_with_style).</param>
    /// <param name="indentLevel">Indentation level (for add_with_style).</param>
    /// <param name="leftIndent">Left indentation in points (for add_with_style).</param>
    /// <param name="firstLineIndent">First line indentation in points (for add_with_style).</param>
    /// <param name="paragraphIndexForAdd">Paragraph index to insert after (for add_with_style).</param>
    /// <param name="tabStops">Custom tab stops as JSON array (for add_with_style).</param>
    /// <returns>A message indicating the result of the operation, or JSON data for search operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "word_text")]
    [Description(
        @"Perform text operations in Word documents. Supports 8 operations: add, delete, replace, search, format, insert_at_position, delete_range, add_with_style.

Usage examples:
- Add text: word_text(operation='add', path='doc.docx', text='Hello World')
- Add formatted text: word_text(operation='add', path='doc.docx', text='Bold text', bold=true)
- Replace text: word_text(operation='replace', path='doc.docx', find='old', replace='new')
- Search text: word_text(operation='search', path='doc.docx', searchText='keyword')
- Format text: word_text(operation='format', path='doc.docx', paragraphIndex=0, runIndex=0, bold=true)
- Insert at position: word_text(operation='insert_at_position', path='doc.docx', paragraphIndex=0, runIndex=0, text='Inserted')
- Delete text: word_text(operation='delete', path='doc.docx', searchText='text to delete') or word_text(operation='delete', path='doc.docx', startParagraphIndex=0, endParagraphIndex=0)
- Delete range: word_text(operation='delete_range', path='doc.docx', startParagraphIndex=0, startRunIndex=0, endParagraphIndex=0, endRunIndex=5)")]
    public string Execute(
        [Description(@"Operation to perform.
- 'add': Add text at document end (required params: path, text)
- 'delete': Delete text (required params: path, searchText OR startParagraphIndex+endParagraphIndex)
- 'replace': Replace text (required params: path, find, replace)
- 'search': Search for text (required params: path, searchText)
- 'format': Format existing text (required params: path, paragraphIndex, runIndex)
- 'insert_at_position': Insert text at specific position (required params: path, paragraphIndex, runIndex, text)
- 'delete_range': Delete text range (required params: path, startParagraphIndex, startRunIndex, endParagraphIndex, endRunIndex)
- 'add_with_style': Add text with style (required params: path, text, styleName)")]
        string operation,
        [Description("Document file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (if not provided, overwrites input, for write operations)")]
        string? outputPath = null,
        // Common parameters
        [Description("Text content (required for add, replace, insert_at_position, add_with_style operations)")]
        string? text = null,
        // Add/AddWithStyle parameters
        [Description("Font name (optional, e.g., 'Arial')")]
        string? fontName = null,
        [Description("Font name for ASCII characters (English, e.g., 'Times New Roman')")]
        string? fontNameAscii = null,
        [Description("Font name for Far East characters (Chinese/Japanese/Korean, optional)")]
        string? fontNameFarEast = null,
        [Description("Font size (optional)")] double? fontSize = null,
        [Description("Bold text (optional)")] bool? bold = null,
        [Description("Italic text (optional)")]
        bool? italic = null,
        [Description("Underline style: none, single, double, dotted, dash (optional, for add/format operations)")]
        string? underline = null,
        [Description(
            "Text color (hex format like 'FF0000' or '#FF0000' for red, or name like 'Red', 'Blue', optional)")]
        string? color = null,
        [Description("Strikethrough (optional, for add/format operations)")]
        bool? strikethrough = null,
        [Description("Superscript (optional, for add/format operations)")]
        bool? superscript = null,
        [Description("Subscript (optional, for add/format operations)")]
        bool? subscript = null,
        // Replace parameters
        [Description("Text to find (required for replace operation)")]
        string? find = null,
        [Description("Replacement text (required for replace operation)")]
        string? replace = null,
        [Description("Use regex matching (optional, for replace/search operations)")]
        bool useRegex = false,
        [Description(
            "Replace text inside fields (optional, default: false). If false, fields like hyperlinks will be excluded from replacement to preserve their functionality")]
        bool replaceInFields = false,
        // Search parameters
        [Description(
            "Text to search for (required for search operation, optional for delete operation as alternative to startParagraphIndex+endParagraphIndex)")]
        string? searchText = null,
        [Description("Case sensitive search (optional, default: false, for search operation)")]
        bool caseSensitive = false,
        [Description("Maximum number of results to return (optional, default: 50, for search operation)")]
        int maxResults = 50,
        [Description(
            "Number of characters to show before and after match for context (optional, default: 50, for search operation)")]
        int contextLength = 50,
        // Delete parameters
        [Description("Start paragraph index (0-based, required for delete operation if searchText is not provided)")]
        int? startParagraphIndex = null,
        [Description("Start run index within start paragraph (0-based, optional, default: 0, for delete operation)")]
        int startRunIndex = 0,
        [Description(
            "End paragraph index (0-based, inclusive, required for delete operation if searchText is not provided)")]
        int? endParagraphIndex = null,
        [Description(
            "End run index within end paragraph (0-based, inclusive, optional, default: last run, for delete operation)")]
        int? endRunIndex = null,
        // Format parameters
        [Description("Paragraph index (0-based, required for format operation)")]
        int? paragraphIndex = null,
        [Description(
            "Run index within the paragraph (0-based, optional, formats all runs if not provided, for format operation)")]
        int? runIndex = null,
        // Insert at position parameters
        [Description("Paragraph index (0-based, required for insert_at_position operation)")]
        int? insertParagraphIndex = null,
        [Description("Character index within paragraph (0-based, required for insert_at_position operation)")]
        int? charIndex = null,
        [Description(
            "Section index (0-based, optional, default: 0, for format/insert_at_position/delete_range operations)")]
        int? sectionIndex = null,
        [Description(
            "Insert before position (optional, default: false, inserts after, for insert_at_position operation)")]
        bool insertBefore = false,
        // Delete range parameters
        [Description("Start character index within paragraph (0-based, required for delete_range operation)")]
        int? startCharIndex = null,
        [Description("End character index within paragraph (0-based, required for delete_range operation)")]
        int? endCharIndex = null,
        // AddWithStyle parameters
        [Description("Style name to apply (e.g., 'Heading 1', 'Normal', optional, for add_with_style operation)")]
        string? styleName = null,
        [Description("Text alignment: left, center, right, justify (optional, for add_with_style operation)")]
        string? alignment = null,
        [Description(
            "Indentation level (0-8, where each level = 36 points / 0.5 inch, optional, for add_with_style operation)")]
        int? indentLevel = null,
        [Description("Left indentation in points (optional, for add_with_style operation)")]
        double? leftIndent = null,
        [Description(
            "First line indentation in points (positive = indent first line, negative = hanging indent, optional, for add_with_style operation)")]
        double? firstLineIndent = null,
        [Description(
            "Index of the paragraph to insert after (0-based, optional, for add_with_style operation). Use -1 to insert at the beginning.")]
        int? paragraphIndexForAdd = null,
        [Description(
            "Custom tab stops as JSON array (optional, for add_with_style operation). Example: [{\"position\":72,\"alignment\":\"Left\",\"leader\":\"None\"}]")]
        string? tabStops = null)
    {
        var effectiveOutputPath = outputPath ?? path;
        if (!string.IsNullOrEmpty(effectiveOutputPath))
            SecurityHelper.ValidateFilePath(effectiveOutputPath, "outputPath", true);

        var parameters = BuildParameters(text, fontName, fontNameAscii, fontNameFarEast, fontSize, bold, italic,
            underline, color, strikethrough, superscript, subscript, find, replace, useRegex, replaceInFields,
            searchText, caseSensitive, maxResults, contextLength, startParagraphIndex, startRunIndex,
            endParagraphIndex, endRunIndex, paragraphIndex, runIndex, insertParagraphIndex, charIndex, sectionIndex,
            insertBefore, startCharIndex, endCharIndex, styleName, alignment, indentLevel, leftIndent,
            firstLineIndent, paragraphIndexForAdd, tabStops);

        var handler = _handlerRegistry.GetHandler(operation);

        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);

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

        return AppendOutputMessage(result, ctx, effectiveOutputPath);
    }

    /// <summary>
    ///     Builds the operation parameters from input values.
    /// </summary>
    private static OperationParameters BuildParameters(
        string? text, string? fontName, string? fontNameAscii, string? fontNameFarEast, double? fontSize,
        bool? bold, bool? italic, string? underline, string? color, bool? strikethrough, bool? superscript,
        bool? subscript, string? find, string? replace, bool useRegex, bool replaceInFields, string? searchText,
        bool caseSensitive, int maxResults, int contextLength, int? startParagraphIndex, int startRunIndex,
        int? endParagraphIndex, int? endRunIndex, int? paragraphIndex, int? runIndex, int? insertParagraphIndex,
        int? charIndex, int? sectionIndex, bool insertBefore, int? startCharIndex, int? endCharIndex,
        string? styleName, string? alignment, int? indentLevel, double? leftIndent, double? firstLineIndent,
        int? paragraphIndexForAdd, string? tabStops)
    {
        var parameters = new OperationParameters();

        // Common parameters
        parameters.Set("text", text);
        parameters.Set("fontName", fontName);
        parameters.Set("fontNameAscii", fontNameAscii);
        parameters.Set("fontNameFarEast", fontNameFarEast);
        parameters.Set("fontSize", fontSize);
        parameters.Set("bold", bold);
        parameters.Set("italic", italic);
        parameters.Set("underline", underline);
        parameters.Set("color", color);
        parameters.Set("strikethrough", strikethrough);
        parameters.Set("superscript", superscript);
        parameters.Set("subscript", subscript);

        // Replace parameters
        parameters.Set("find", find);
        parameters.Set("replace", replace);
        parameters.Set("useRegex", useRegex);
        parameters.Set("replaceInFields", replaceInFields);

        // Search parameters
        parameters.Set("searchText", searchText);
        parameters.Set("caseSensitive", caseSensitive);
        parameters.Set("maxResults", maxResults);
        parameters.Set("contextLength", contextLength);

        // Delete parameters
        parameters.Set("startParagraphIndex", startParagraphIndex);
        parameters.Set("startRunIndex", startRunIndex);
        parameters.Set("endParagraphIndex", endParagraphIndex);
        parameters.Set("endRunIndex", endRunIndex);

        // Format parameters
        parameters.Set("paragraphIndex", paragraphIndex);
        parameters.Set("runIndex", runIndex);
        parameters.Set("sectionIndex", sectionIndex);

        // Insert at position parameters
        parameters.Set("insertParagraphIndex", insertParagraphIndex);
        parameters.Set("charIndex", charIndex);
        parameters.Set("insertBefore", insertBefore);

        // Delete range parameters
        parameters.Set("startCharIndex", startCharIndex);
        parameters.Set("endCharIndex", endCharIndex);

        // AddWithStyle parameters
        parameters.Set("styleName", styleName);
        parameters.Set("alignment", alignment);
        parameters.Set("indentLevel", indentLevel);
        parameters.Set("leftIndent", leftIndent);
        parameters.Set("firstLineIndent", firstLineIndent);
        parameters.Set("paragraphIndexForAdd", paragraphIndexForAdd);
        parameters.Set("tabStops", tabStops);

        return parameters;
    }

    /// <summary>
    ///     Appends output message to the result if not already present.
    /// </summary>
    /// <param name="result">The handler result.</param>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <returns>Result with output message appended.</returns>
    private static string AppendOutputMessage(string result, DocumentContext<Document> ctx, string? outputPath)
    {
        var outputMessage = ctx.GetOutputMessage(outputPath);
        if (!result.Contains(outputMessage))
            return $"{result}\n{outputMessage}";
        return result;
    }
}
