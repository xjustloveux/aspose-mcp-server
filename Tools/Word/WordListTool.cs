using System.ComponentModel;
using System.Text.Json.Nodes;
using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Unified tool for list operations in Word documents
///     Merges: WordAddListTool, WordAddListItemTool, WordDeleteListItemTool, WordEditListItemTool,
///     WordSetListFormatTool, WordGetListFormatTool
/// </summary>
[McpServerToolType]
public class WordListTool
{
    /// <summary>
    ///     Handler registry for list operations
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
    ///     Initializes a new instance of the WordListTool class
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document operations</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation</param>
    public WordListTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry = HandlerRegistry<Document>.CreateFromNamespace("AsposeMcpServer.Handlers.Word.List");
    }

    /// <summary>
    ///     Executes a Word list operation (add_list, add_item, delete_item, edit_item, set_format, get_format,
    ///     restart_numbering, convert_to_list).
    /// </summary>
    /// <param name="operation">
    ///     The operation to perform: add_list, add_item, delete_item, edit_item, set_format, get_format,
    ///     restart_numbering, convert_to_list.
    /// </param>
    /// <param name="path">Word document file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="items">List items for add_list operation (string array or object array with text/level).</param>
    /// <param name="listType">List type: bullet, number, custom (default: bullet).</param>
    /// <param name="bulletChar">Custom bullet character (for custom type).</param>
    /// <param name="numberFormat">Number format: arabic, roman, letter (default: arabic).</param>
    /// <param name="continuePrevious">Continue numbering from last list (default: false).</param>
    /// <param name="text">List item text content.</param>
    /// <param name="styleName">Style name for the list item.</param>
    /// <param name="listLevel">List level (0-8).</param>
    /// <param name="applyStyleIndent">Use style-defined indent (default: true).</param>
    /// <param name="paragraphIndex">Paragraph index (0-based).</param>
    /// <param name="level">List level for edit (0-8).</param>
    /// <param name="numberStyle">Number style: arabic, roman, letter, bullet, none.</param>
    /// <param name="indentLevel">Indentation level (0-8, each level = 36 points).</param>
    /// <param name="leftIndent">Left indent in points.</param>
    /// <param name="firstLineIndent">First line indent in points.</param>
    /// <param name="startAt">Number to restart at (default: 1).</param>
    /// <param name="startParagraphIndex">Starting paragraph index (for convert_to_list).</param>
    /// <param name="endParagraphIndex">Ending paragraph index (for convert_to_list).</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get_format operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "word_list")]
    [Description(
        @"Manage lists in Word documents. Supports 8 operations: add_list, add_item, delete_item, edit_item, set_format, get_format, restart_numbering, convert_to_list.

Usage examples:
- Add bullet list: word_list(operation='add_list', path='doc.docx', items=['Item 1', 'Item 2', 'Item 3'])
- Add numbered list: word_list(operation='add_list', path='doc.docx', items=['First', 'Second'], listType='number')
- Add list item: word_list(operation='add_item', path='doc.docx', text='New item', styleName='Heading 4')
- Delete list item: word_list(operation='delete_item', path='doc.docx', paragraphIndex=0)
- Edit list item: word_list(operation='edit_item', path='doc.docx', paragraphIndex=0, text='Updated text')
- Get list format: word_list(operation='get_format', path='doc.docx', paragraphIndex=0)
- Restart numbering: word_list(operation='restart_numbering', path='doc.docx', paragraphIndex=2, startAt=1)
- Convert to list: word_list(operation='convert_to_list', path='doc.docx', startParagraphIndex=0, endParagraphIndex=5)")]
    public string Execute( // NOSONAR S107 - MCP protocol requires multiple parameters
        [Description(
            "Operation: add_list, add_item, delete_item, edit_item, set_format, get_format, restart_numbering, convert_to_list")]
        string operation,
        [Description("Document file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("List items for add_list operation (string array or object array with text/level)")]
        JsonArray? items = null,
        [Description("List type: bullet, number, custom (default: bullet)")]
        string listType = "bullet",
        [Description("Custom bullet character (for custom type)")]
        string bulletChar = "•",
        [Description("Number format: arabic, roman, letter (default: arabic)")]
        string numberFormat = "arabic",
        [Description("Continue numbering from last list (default: false)")]
        bool continuePrevious = false,
        [Description("List item text content")]
        string? text = null,
        [Description("Style name for the list item")]
        string? styleName = null,
        [Description("List level (0-8)")] int listLevel = 0,
        [Description("Use style-defined indent (default: true)")]
        bool applyStyleIndent = true,
        [Description("Paragraph index (0-based)")]
        int? paragraphIndex = null,
        [Description("List level for edit (0-8)")]
        int? level = null,
        [Description("Number style: arabic, roman, letter, bullet, none")]
        string? numberStyle = null,
        [Description("Indentation level (0-8, each level = 36 points)")]
        int? indentLevel = null,
        [Description("Left indent in points")] double? leftIndent = null,
        [Description("First line indent in points")]
        double? firstLineIndent = null,
        [Description("Number to restart at (default: 1)")]
        int startAt = 1,
        [Description("Starting paragraph index (for convert_to_list)")]
        int? startParagraphIndex = null,
        [Description("Ending paragraph index (for convert_to_list)")]
        int? endParagraphIndex = null)
    {
        var parameters = BuildParameters(operation, items, listType, bulletChar, numberFormat, continuePrevious,
            text, styleName, listLevel, applyStyleIndent, paragraphIndex, level, numberStyle, indentLevel,
            leftIndent, firstLineIndent, startAt, startParagraphIndex, endParagraphIndex);

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
    private static OperationParameters BuildParameters( // NOSONAR S107
        string operation,
        JsonArray? items,
        string listType,
        string bulletChar,
        string numberFormat,
        bool continuePrevious,
        string? text,
        string? styleName,
        int listLevel,
        bool applyStyleIndent,
        int? paragraphIndex,
        int? level,
        string? numberStyle,
        int? indentLevel,
        double? leftIndent,
        double? firstLineIndent,
        int startAt,
        int? startParagraphIndex,
        int? endParagraphIndex)
    {
        var parameters = new OperationParameters();

        return operation.ToLower() switch
        {
            "add_list" => BuildAddListParameters(parameters, items, listType, bulletChar, numberFormat,
                continuePrevious),
            "add_item" => BuildAddItemParameters(parameters, text, styleName, listLevel, applyStyleIndent),
            "delete_item" or "get_format" => BuildParagraphIndexParameters(parameters, paragraphIndex),
            "edit_item" => BuildEditItemParameters(parameters, paragraphIndex, text, level),
            "set_format" => BuildSetFormatParameters(parameters, paragraphIndex, numberStyle, indentLevel, leftIndent,
                firstLineIndent),
            "restart_numbering" => BuildRestartNumberingParameters(parameters, paragraphIndex, startAt),
            "convert_to_list" => BuildConvertToListParameters(parameters, startParagraphIndex, endParagraphIndex,
                listType, numberFormat),
            _ => parameters
        };
    }

    /// <summary>
    ///     Builds parameters for the add list operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="items">The list items (string array or object array with text/level).</param>
    /// <param name="listType">The list type: 'bullet', 'number', 'custom'.</param>
    /// <param name="bulletChar">The custom bullet character.</param>
    /// <param name="numberFormat">The number format: 'arabic', 'roman', 'letter'.</param>
    /// <param name="continuePrevious">Whether to continue numbering from last list.</param>
    /// <returns>OperationParameters configured for the add list operation.</returns>
    private static OperationParameters BuildAddListParameters(OperationParameters parameters, JsonArray? items,
        string listType, string bulletChar, string numberFormat, bool continuePrevious)
    {
        if (items != null) parameters.Set("items", items);
        parameters.Set("listType", listType);
        parameters.Set("bulletChar", bulletChar);
        parameters.Set("numberFormat", numberFormat);
        parameters.Set("continuePrevious", continuePrevious);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the add list item operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="text">The list item text content.</param>
    /// <param name="styleName">The style name for the list item.</param>
    /// <param name="listLevel">The list level (0-8).</param>
    /// <param name="applyStyleIndent">Whether to use style-defined indent.</param>
    /// <returns>OperationParameters configured for the add item operation.</returns>
    private static OperationParameters BuildAddItemParameters(OperationParameters parameters, string? text,
        string? styleName, int listLevel, bool applyStyleIndent)
    {
        if (text != null) parameters.Set("text", text);
        if (styleName != null) parameters.Set("styleName", styleName);
        parameters.Set("listLevel", listLevel);
        parameters.Set("applyStyleIndent", applyStyleIndent);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for paragraph index-based operations (delete_item, get_format).
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="paragraphIndex">The paragraph index (0-based).</param>
    /// <returns>OperationParameters configured for paragraph index-based operations.</returns>
    private static OperationParameters BuildParagraphIndexParameters(OperationParameters parameters,
        int? paragraphIndex)
    {
        if (paragraphIndex.HasValue) parameters.Set("paragraphIndex", paragraphIndex.Value);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the edit list item operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="paragraphIndex">The paragraph index (0-based).</param>
    /// <param name="text">The updated text content.</param>
    /// <param name="level">The list level (0-8).</param>
    /// <returns>OperationParameters configured for the edit item operation.</returns>
    private static OperationParameters BuildEditItemParameters(OperationParameters parameters, int? paragraphIndex,
        string? text, int? level)
    {
        if (paragraphIndex.HasValue) parameters.Set("paragraphIndex", paragraphIndex.Value);
        if (text != null) parameters.Set("text", text);
        if (level.HasValue) parameters.Set("level", level.Value);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the set list format operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="paragraphIndex">The paragraph index (0-based).</param>
    /// <param name="numberStyle">The number style: 'arabic', 'roman', 'letter', 'bullet', 'none'.</param>
    /// <param name="indentLevel">The indentation level (0-8, each level = 36 points).</param>
    /// <param name="leftIndent">The left indent in points.</param>
    /// <param name="firstLineIndent">The first line indent in points.</param>
    /// <returns>OperationParameters configured for the set format operation.</returns>
    private static OperationParameters BuildSetFormatParameters(OperationParameters parameters, int? paragraphIndex,
        string? numberStyle, int? indentLevel, double? leftIndent, double? firstLineIndent)
    {
        if (paragraphIndex.HasValue) parameters.Set("paragraphIndex", paragraphIndex.Value);
        if (numberStyle != null) parameters.Set("numberStyle", numberStyle);
        if (indentLevel.HasValue) parameters.Set("indentLevel", indentLevel.Value);
        if (leftIndent.HasValue) parameters.Set("leftIndent", leftIndent.Value);
        if (firstLineIndent.HasValue) parameters.Set("firstLineIndent", firstLineIndent.Value);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the restart numbering operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="paragraphIndex">The paragraph index (0-based).</param>
    /// <param name="startAt">The number to restart at.</param>
    /// <returns>OperationParameters configured for the restart numbering operation.</returns>
    private static OperationParameters BuildRestartNumberingParameters(OperationParameters parameters,
        int? paragraphIndex, int startAt)
    {
        if (paragraphIndex.HasValue) parameters.Set("paragraphIndex", paragraphIndex.Value);
        parameters.Set("startAt", startAt);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the convert to list operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="startParagraphIndex">The starting paragraph index.</param>
    /// <param name="endParagraphIndex">The ending paragraph index.</param>
    /// <param name="listType">The list type: 'bullet', 'number', 'custom'.</param>
    /// <param name="numberFormat">The number format: 'arabic', 'roman', 'letter'.</param>
    /// <returns>OperationParameters configured for the convert to list operation.</returns>
    private static OperationParameters BuildConvertToListParameters(OperationParameters parameters,
        int? startParagraphIndex, int? endParagraphIndex, string listType, string numberFormat)
    {
        if (startParagraphIndex.HasValue) parameters.Set("startParagraphIndex", startParagraphIndex.Value);
        if (endParagraphIndex.HasValue) parameters.Set("endParagraphIndex", endParagraphIndex.Value);
        parameters.Set("listType", listType);
        parameters.Set("numberFormat", numberFormat);
        return parameters;
    }
}
