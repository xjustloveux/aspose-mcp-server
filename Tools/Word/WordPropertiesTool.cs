using System.ComponentModel;
using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Core.Tools;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Unified tool for managing Word document properties (get, set)
///     Merges: WordGetDocumentPropertiesTool, WordSetDocumentPropertiesTool, WordSetPropertiesTool
/// </summary>
[ToolHandlerMapping("AsposeMcpServer.Handlers.Word.Properties")]
[McpServerToolType]
public class WordPropertiesTool : PropertiesToolBase<Document>
{
    /// <summary>
    ///     Initializes a new instance of the WordPropertiesTool class
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document operations</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation</param>
    public WordPropertiesTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
        : base(sessionManager, identityAccessor, "AsposeMcpServer.Handlers.Word.Properties")
    {
    }

    /// <summary>
    ///     Executes a Word properties operation (get, set).
    /// </summary>
    /// <param name="operation">The operation to perform: get, set.</param>
    /// <param name="path">Word document file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (for set operation).</param>
    /// <param name="title">Document title (for set).</param>
    /// <param name="subject">Document subject (for set).</param>
    /// <param name="author">Document author (for set).</param>
    /// <param name="keywords">Keywords (for set).</param>
    /// <param name="comments">Comments (for set).</param>
    /// <param name="category">Document category (for set).</param>
    /// <param name="company">Company name (for set).</param>
    /// <param name="manager">Manager name (for set).</param>
    /// <param name="customProperties">Custom properties as JSON string (for set).</param>
    /// <returns>Document properties as JSON for get, or a success message for set operations.</returns>
    /// <exception cref="ArgumentException">Thrown when the operation is unknown.</exception>
    [McpServerTool(
        Name = "word_properties",
        Title = "Word Properties Operations",
        Destructive = false,
        Idempotent = true,
        OpenWorld = false,
        ReadOnly = false,
        UseStructuredContent = true)]
    [Description(@"Get or set Word document properties (metadata). Supports 2 operations: get, set.

Usage examples:
- Get properties: word_properties(operation='get', path='doc.docx')
- Set properties: word_properties(operation='set', path='doc.docx', title='Title', author='Author', subject='Subject')

Notes:
- The 'set' operation is for content metadata (title, author, subject, etc.), not for statistics (word count, page count)
- Statistics like word count and page count are automatically calculated by Word and cannot be manually set
- Custom properties support multiple types: string, number (integer/double), boolean, and datetime (ISO 8601 format)")]
    public object Execute(
        [Description("Operation: get, set")] string operation,
        [Description("Document file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (if not provided, overwrites input, for set operation)")]
        string? outputPath = null,
        [Description("Document title (optional, for set operation)")]
        string? title = null,
        [Description("Document subject (optional, for set operation)")]
        string? subject = null,
        [Description("Document author (optional, for set operation)")]
        string? author = null,
        [Description("Keywords (optional, for set operation)")]
        string? keywords = null,
        [Description("Comments (optional, for set operation)")]
        string? comments = null,
        [Description("Document category (optional, for set operation)")]
        string? category = null,
        [Description("Company name (optional, for set operation)")]
        string? company = null,
        [Description("Manager name (optional, for set operation)")]
        string? manager = null,
        [Description(
            "Custom properties as JSON string (optional, for set operation). Supports string, number (integer/double), boolean, and datetime (ISO 8601 format).")]
        string? customProperties = null)
    {
        var parameters = BuildParameters(operation, title, subject, author, keywords, comments, category, company,
            manager, customProperties);

        return ExecuteOperation(
            operation,
            sessionId,
            path,
            outputPath,
            parameters);
    }

    /// <summary>
    ///     Builds OperationParameters from method parameters.
    /// </summary>
    /// <param name="operation">The operation to perform.</param>
    /// <param name="title">The document title.</param>
    /// <param name="subject">The document subject.</param>
    /// <param name="author">The document author.</param>
    /// <param name="keywords">The document keywords.</param>
    /// <param name="comments">The document comments.</param>
    /// <param name="category">The document category.</param>
    /// <param name="company">The company name.</param>
    /// <param name="manager">The manager name.</param>
    /// <param name="customProperties">Custom properties as JSON string.</param>
    /// <returns>OperationParameters configured for the properties operation.</returns>
    private static OperationParameters BuildParameters(
        string operation,
        string? title,
        string? subject,
        string? author,
        string? keywords,
        string? comments,
        string? category,
        string? company,
        string? manager,
        string? customProperties)
    {
        var parameters = new OperationParameters();

        if (string.Equals(operation, "set", StringComparison.OrdinalIgnoreCase))
        {
            if (title != null) parameters.Set("title", title);
            if (subject != null) parameters.Set("subject", subject);
            if (author != null) parameters.Set("author", author);
            if (keywords != null) parameters.Set("keywords", keywords);
            if (comments != null) parameters.Set("comments", comments);
            if (category != null) parameters.Set("category", category);
            if (company != null) parameters.Set("company", company);
            if (manager != null) parameters.Set("manager", manager);
            if (customProperties != null) parameters.Set("customProperties", customProperties);
        }

        return parameters;
    }
}
