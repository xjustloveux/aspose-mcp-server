using System.ComponentModel;
using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.PowerPoint.Properties;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for managing PowerPoint document properties (get, set).
///     Uses IPresentationInfo for efficient property reading without loading entire presentation.
/// </summary>
[ToolHandlerMapping("AsposeMcpServer.Handlers.PowerPoint.Properties")]
[McpServerToolType]
public class PptPropertiesTool
{
    /// <summary>
    ///     Handler registry for properties operations.
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
    ///     Initializes a new instance of the <see cref="PptPropertiesTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory editing.</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation.</param>
    public PptPropertiesTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry =
            HandlerRegistry<Presentation>.CreateFromNamespace("AsposeMcpServer.Handlers.PowerPoint.Properties");
    }

    /// <summary>
    ///     Executes a PowerPoint properties operation (get, set).
    /// </summary>
    /// <param name="operation">The operation to perform: get, set.</param>
    /// <param name="path">Presentation file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="title">Title (optional, for set).</param>
    /// <param name="subject">Subject (optional, for set).</param>
    /// <param name="author">Author (optional, for set).</param>
    /// <param name="keywords">Keywords (optional, for set).</param>
    /// <param name="comments">Comments (optional, for set).</param>
    /// <param name="category">Category (optional, for set).</param>
    /// <param name="company">Company (optional, for set).</param>
    /// <param name="manager">Manager (optional, for set).</param>
    /// <param name="customProperties">
    ///     Custom properties as key-value pairs. Supports: string, int, double, bool, DateTime (ISO
    ///     format).
    /// </param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(
        Name = "ppt_properties",
        Title = "PowerPoint Properties Operations",
        Destructive = false,
        Idempotent = true,
        OpenWorld = false,
        ReadOnly = false,
        UseStructuredContent = true)]
    [Description(@"Manage PowerPoint document properties. Supports 2 operations: get, set.

Warning: If outputPath is not provided for 'set' operation, the original file will be overwritten.
Note: Custom properties support multiple types: string, int, double, bool, DateTime (ISO format).

Usage examples:
- Get properties: ppt_properties(operation='get', path='presentation.pptx')
- Set properties: ppt_properties(operation='set', path='presentation.pptx', title='Title', author='Author')
- Set custom properties: ppt_properties(operation='set', path='presentation.pptx', customProperties={'Count': 42, 'IsPublished': true})")]
    public object Execute(
        [Description("Operation: get, set")] string operation,
        [Description("Presentation file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Title (optional, for set)")]
        string? title = null,
        [Description("Subject (optional, for set)")]
        string? subject = null,
        [Description("Author (optional, for set)")]
        string? author = null,
        [Description("Keywords (optional, for set)")]
        string? keywords = null,
        [Description("Comments (optional, for set)")]
        string? comments = null,
        [Description("Category (optional, for set)")]
        string? category = null,
        [Description("Company (optional, for set)")]
        string? company = null,
        [Description("Manager (optional, for set)")]
        string? manager = null,
        [Description(
            "Custom properties as key-value pairs. Supports: string, int, double, bool, DateTime (ISO format).")]
        Dictionary<string, object>? customProperties = null)
    {
        if (operation.Equals("get", StringComparison.OrdinalIgnoreCase) && string.IsNullOrEmpty(sessionId))
            return ResultHelper.FinalizeResult((dynamic)GetPropertiesEfficient(path), outputPath, sessionId);

        using var ctx = DocumentContext<Presentation>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters = BuildParameters(title, subject, author, keywords, comments,
            category, company, manager, customProperties);

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

        if (operation.Equals("get", StringComparison.OrdinalIgnoreCase))
            return ResultHelper.FinalizeResult((dynamic)result, ctx, outputPath);

        if (operationContext.IsModified)
            ctx.Save(outputPath);

        return ResultHelper.FinalizeResult((dynamic)result, ctx, outputPath);
    }

    /// <summary>
    ///     Gets presentation properties using IPresentationInfo for efficiency.
    ///     This method reads properties without loading the entire presentation.
    /// </summary>
    /// <param name="path">The presentation file path.</param>
    /// <returns>An object containing the document properties.</returns>
    /// <exception cref="ArgumentException">Thrown when path is not provided.</exception>
    private static GetPropertiesPptResult GetPropertiesEfficient(string? path)
    {
        if (string.IsNullOrEmpty(path))
            throw new ArgumentException("Either sessionId or path must be provided");

        SecurityHelper.ValidateFilePath(path, nameof(path), true);

        var info = PresentationFactory.Instance.GetPresentationInfo(path);
        var props = info.ReadDocumentProperties();

        return new GetPropertiesPptResult
        {
            Title = props.Title,
            Subject = props.Subject,
            Author = props.Author,
            Keywords = props.Keywords,
            Comments = props.Comments,
            Category = props.Category,
            Company = props.Company,
            Manager = props.Manager,
            CreatedTime = props.CreatedTime,
            LastSavedTime = props.LastSavedTime,
            RevisionNumber = props.RevisionNumber
        };
    }

    /// <summary>
    ///     Builds OperationParameters from method parameters.
    /// </summary>
    /// <param name="title">The presentation title.</param>
    /// <param name="subject">The presentation subject.</param>
    /// <param name="author">The presentation author.</param>
    /// <param name="keywords">The presentation keywords.</param>
    /// <param name="comments">The presentation comments.</param>
    /// <param name="category">The presentation category.</param>
    /// <param name="company">The company name.</param>
    /// <param name="manager">The manager name.</param>
    /// <param name="customProperties">Custom properties as key-value pairs.</param>
    /// <returns>OperationParameters configured for the properties operation.</returns>
    private static OperationParameters BuildParameters(
        string? title,
        string? subject,
        string? author,
        string? keywords,
        string? comments,
        string? category,
        string? company,
        string? manager,
        Dictionary<string, object>? customProperties)
    {
        var parameters = new OperationParameters();

        if (title != null) parameters.Set("title", title);
        if (subject != null) parameters.Set("subject", subject);
        if (author != null) parameters.Set("author", author);
        if (keywords != null) parameters.Set("keywords", keywords);
        if (comments != null) parameters.Set("comments", comments);
        if (category != null) parameters.Set("category", category);
        if (company != null) parameters.Set("company", company);
        if (manager != null) parameters.Set("manager", manager);
        if (customProperties != null) parameters.Set("customProperties", customProperties);

        return parameters;
    }
}
