using System.ComponentModel;
using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Helpers;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Tool for managing content controls (structured document tags) in Word documents (add, edit, delete, get,
///     set_value).
/// </summary>
[ToolHandlerMapping("AsposeMcpServer.Handlers.Word.ContentControl")]
[McpServerToolType]
public class WordContentControlTool
{
    /// <summary>
    ///     Handler registry for content control operations.
    /// </summary>
    private readonly HandlerRegistry<Document> _handlerRegistry;

    /// <summary>
    ///     The session identity accessor for session isolation.
    /// </summary>
    private readonly ISessionIdentityAccessor? _identityAccessor;

    /// <summary>
    ///     The document session manager for managing in-memory document sessions.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="WordContentControlTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document operations.</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation.</param>
    public WordContentControlTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry =
            HandlerRegistry<Document>.CreateFromNamespace("AsposeMcpServer.Handlers.Word.ContentControl");
    }

    /// <summary>
    ///     Executes a content control operation on a Word document (add, edit, delete, get, set_value).
    /// </summary>
    /// <param name="operation">The operation to perform: add, edit, delete, get, set_value.</param>
    /// <param name="path">Word document file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="type">
    ///     Content control type for add: PlainText, RichText, DropDownList, DatePicker, CheckBox, Picture,
    ///     ComboBox.
    /// </param>
    /// <param name="tag">Tag identifier for the content control.</param>
    /// <param name="title">Display title for the content control.</param>
    /// <param name="value">Value for the content control (for add, set_value).</param>
    /// <param name="items">Comma-separated items for DropDownList/ComboBox (for add).</param>
    /// <param name="index">Zero-based index of the content control (for edit, delete, set_value).</param>
    /// <param name="newTag">New tag value (for edit).</param>
    /// <param name="newTitle">New title value (for edit).</param>
    /// <param name="lockContents">Lock content control contents (for add, edit).</param>
    /// <param name="lockDeletion">Lock content control from deletion (for add, edit).</param>
    /// <param name="keepContent">Keep content when deleting (default: true).</param>
    /// <param name="paragraphIndex">Paragraph index for insertion (0-based, for add).</param>
    /// <returns>A message or data indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(
        Name = "word_content_control",
        Title = "Word Content Control Operations",
        Destructive = true,
        Idempotent = false,
        OpenWorld = false,
        ReadOnly = false,
        UseStructuredContent = true)]
    [Description(
        @"Manage content controls (structured document tags) in Word documents. Supports 5 operations: add, edit, delete, get, set_value.

Usage examples:
- Add plain text control: word_content_control(operation='add', path='doc.docx', type='PlainText', tag='name', title='Full Name')
- Add dropdown: word_content_control(operation='add', path='doc.docx', type='DropDownList', tag='color', items='Red,Green,Blue')
- Add checkbox: word_content_control(operation='add', path='doc.docx', type='CheckBox', tag='agree', value='false')
- Get all controls: word_content_control(operation='get', path='doc.docx')
- Set value by tag: word_content_control(operation='set_value', path='doc.docx', tag='name', value='John Doe')
- Edit properties: word_content_control(operation='edit', path='doc.docx', index=0, newTag='updated', lockContents=true)
- Delete control: word_content_control(operation='delete', path='doc.docx', tag='name', keepContent=true)")]
    public object Execute(
        [Description(@"Operation to perform.
- 'add': Add a new content control (required params: type)
- 'edit': Edit content control properties (required params: index or tag)
- 'delete': Delete a content control (required params: index or tag)
- 'get': Get all content controls
- 'set_value': Set the value of a content control (required params: index or tag, value)")]
        string operation,
        [Description("Document file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description(
            "Content control type: PlainText, RichText, DropDownList, DatePicker, CheckBox, Picture, ComboBox (for add)")]
        string? type = null,
        [Description("Tag identifier for the content control")]
        string? tag = null,
        [Description("Display title for the content control")]
        string? title = null,
        [Description("Value for the content control")]
        string? value = null,
        [Description("Comma-separated items for DropDownList/ComboBox (for add)")]
        string? items = null,
        [Description("Zero-based index of the content control (for edit, delete, set_value)")]
        int? index = null,
        [Description("New tag value (for edit)")]
        string? newTag = null,
        [Description("New title value (for edit)")]
        string? newTitle = null,
        [Description("Lock content control contents (for add, edit)")]
        bool? lockContents = null,
        [Description("Lock content control from deletion (for add, edit)")]
        bool? lockDeletion = null,
        [Description("Keep content when deleting (default: true)")]
        bool keepContent = true,
        [Description("Paragraph index for insertion (0-based, for add)")]
        int? paragraphIndex = null)
    {
        var parameters = BuildParameters(operation, type, tag, title, value, items, index,
            newTag, newTitle, lockContents, lockDeletion, keepContent, paragraphIndex);

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
        string? type,
        string? tag,
        string? title,
        string? value,
        string? items,
        int? index,
        string? newTag,
        string? newTitle,
        bool? lockContents,
        bool? lockDeletion,
        bool keepContent,
        int? paragraphIndex)
    {
        return operation.ToLowerInvariant() switch
        {
            "add" => BuildAddParameters(type, tag, title, value, items, lockContents, lockDeletion, paragraphIndex),
            "edit" => BuildEditParameters(index, tag, newTag, newTitle, lockContents, lockDeletion),
            "delete" => BuildDeleteParameters(index, tag, keepContent),
            "get" => BuildGetParameters(tag, type),
            "set_value" => BuildSetValueParameters(index, tag, value),
            _ => new OperationParameters()
        };
    }

    /// <summary>
    ///     Builds parameters for the add operation.
    /// </summary>
    /// <param name="type">The content control type.</param>
    /// <param name="tag">The tag identifier.</param>
    /// <param name="title">The display title.</param>
    /// <param name="value">The initial value.</param>
    /// <param name="items">Comma-separated items for DropDownList/ComboBox.</param>
    /// <param name="lockContents">Whether to lock contents.</param>
    /// <param name="lockDeletion">Whether to lock deletion.</param>
    /// <param name="paragraphIndex">The paragraph index to insert at.</param>
    /// <returns>OperationParameters configured for adding a content control.</returns>
    private static OperationParameters BuildAddParameters(string? type, string? tag, string? title,
        string? value, string? items, bool? lockContents, bool? lockDeletion, int? paragraphIndex)
    {
        var parameters = new OperationParameters();
        if (type != null) parameters.Set("type", type);
        if (tag != null) parameters.Set("tag", tag);
        if (title != null) parameters.Set("title", title);
        if (value != null) parameters.Set("value", value);
        if (items != null) parameters.Set("items", items);
        if (lockContents.HasValue) parameters.Set("lockContents", lockContents.Value);
        if (lockDeletion.HasValue) parameters.Set("lockDeletion", lockDeletion.Value);
        if (paragraphIndex.HasValue) parameters.Set("paragraphIndex", paragraphIndex.Value);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the edit operation.
    /// </summary>
    /// <param name="index">The content control index.</param>
    /// <param name="tag">The tag to identify the content control.</param>
    /// <param name="newTag">The new tag value.</param>
    /// <param name="newTitle">The new title value.</param>
    /// <param name="lockContents">Whether to lock contents.</param>
    /// <param name="lockDeletion">Whether to lock deletion.</param>
    /// <returns>OperationParameters configured for editing a content control.</returns>
    private static OperationParameters BuildEditParameters(int? index, string? tag, string? newTag,
        string? newTitle, bool? lockContents, bool? lockDeletion)
    {
        var parameters = new OperationParameters();
        if (index.HasValue) parameters.Set("index", index.Value);
        if (tag != null) parameters.Set("tag", tag);
        if (newTag != null) parameters.Set("newTag", newTag);
        if (newTitle != null) parameters.Set("newTitle", newTitle);
        if (lockContents.HasValue) parameters.Set("lockContents", lockContents.Value);
        if (lockDeletion.HasValue) parameters.Set("lockDeletion", lockDeletion.Value);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the delete operation.
    /// </summary>
    /// <param name="index">The content control index.</param>
    /// <param name="tag">The tag to identify the content control.</param>
    /// <param name="keepContent">Whether to keep content after deletion.</param>
    /// <returns>OperationParameters configured for deleting a content control.</returns>
    private static OperationParameters BuildDeleteParameters(int? index, string? tag, bool keepContent)
    {
        var parameters = new OperationParameters();
        if (index.HasValue) parameters.Set("index", index.Value);
        if (tag != null) parameters.Set("tag", tag);
        parameters.Set("keepContent", keepContent);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the get operation.
    /// </summary>
    /// <param name="tag">Optional tag filter.</param>
    /// <param name="type">Optional type filter.</param>
    /// <returns>OperationParameters configured for getting content controls.</returns>
    private static OperationParameters BuildGetParameters(string? tag, string? type)
    {
        var parameters = new OperationParameters();
        if (tag != null) parameters.Set("tag", tag);
        if (type != null) parameters.Set("type", type);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the set_value operation.
    /// </summary>
    /// <param name="index">The content control index.</param>
    /// <param name="tag">The tag to identify the content control.</param>
    /// <param name="value">The value to set.</param>
    /// <returns>OperationParameters configured for setting the value of a content control.</returns>
    private static OperationParameters BuildSetValueParameters(int? index, string? tag, string? value)
    {
        var parameters = new OperationParameters();
        if (index.HasValue) parameters.Set("index", index.Value);
        if (tag != null) parameters.Set("tag", tag);
        if (value != null) parameters.Set("value", value);
        return parameters;
    }
}
