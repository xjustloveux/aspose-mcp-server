using System.ComponentModel;
using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Tool for Word file operations (create, create_from_template, convert, merge, split).
/// </summary>
[McpServerToolType]
public class WordFileTool
{
    /// <summary>
    ///     Handler registry for file operations
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
    ///     Initializes a new instance of the <see cref="WordFileTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document operations.</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation.</param>
    public WordFileTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry = HandlerRegistry<Document>.CreateFromNamespace("AsposeMcpServer.Handlers.Word.File");
    }

    /// <summary>
    ///     Executes a Word file operation (create, create_from_template, convert, merge, or split).
    /// </summary>
    /// <param name="operation">The operation to perform: create, create_from_template, convert, merge, or split.</param>
    /// <param name="sessionId">Session ID to read document from session (for convert, split, create_from_template).</param>
    /// <param name="path">Input file path (for convert, split).</param>
    /// <param name="outputPath">Output file path (for create, create_from_template, convert, merge).</param>
    /// <param name="templatePath">Template file path (for create_from_template).</param>
    /// <param name="dataJson">JSON data for template rendering (for create_from_template).</param>
    /// <param name="format">Output format: pdf, html, docx, txt, rtf, odt, epub, xps (for convert).</param>
    /// <param name="inputPaths">Array of input file paths to merge (for merge).</param>
    /// <param name="importFormatMode">
    ///     Format mode when merging: KeepSourceFormatting, UseDestinationStyles,
    ///     KeepDifferentStyles.
    /// </param>
    /// <param name="unlinkHeadersFooters">Unlink headers/footers after merge.</param>
    /// <param name="outputDir">Output directory for split files (for split).</param>
    /// <param name="splitBy">Split by: section, page.</param>
    /// <param name="content">Initial content (for create).</param>
    /// <param name="skipInitialContent">Create blank document (for create).</param>
    /// <param name="marginTop">Top margin in points.</param>
    /// <param name="marginBottom">Bottom margin in points.</param>
    /// <param name="marginLeft">Left margin in points.</param>
    /// <param name="marginRight">Right margin in points.</param>
    /// <param name="compatibilityMode">Word compatibility mode.</param>
    /// <param name="paperSize">Predefined paper size.</param>
    /// <param name="pageWidth">Page width in points (overrides paperSize).</param>
    /// <param name="pageHeight">Page height in points (overrides paperSize).</param>
    /// <param name="headerDistance">Header distance from page top in points.</param>
    /// <param name="footerDistance">Footer distance from page bottom in points.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when the operation is unknown or required parameters are missing.</exception>
    /// <exception cref="InvalidOperationException">Thrown when session management is not enabled but sessionId is provided.</exception>
    /// <exception cref="FileNotFoundException">Thrown when the template file is not found.</exception>
    [McpServerTool(Name = "word_file")]
    [Description(
        @"Perform file operations on Word documents. Supports 5 operations: create, create_from_template, convert, merge, split.

Usage examples:
- Create document: word_file(operation='create', outputPath='new.docx')
- Create from template: word_file(operation='create_from_template', templatePath='template.docx', outputPath='output.docx', dataJson='{""Name"":""John""}')
- Create from session template: word_file(operation='create_from_template', sessionId='sess_xxx', outputPath='output.docx', dataJson='{""Name"":""John""}')
- Convert format: word_file(operation='convert', path='doc.docx', outputPath='doc.pdf', format='pdf')
- Convert from session: word_file(operation='convert', sessionId='sess_xxx', outputPath='doc.pdf', format='pdf')
- Merge documents: word_file(operation='merge', inputPaths=['doc1.docx','doc2.docx'], outputPath='merged.docx')
- Split document: word_file(operation='split', path='doc.docx', outputDir='output/', splitBy='page')
- Split from session: word_file(operation='split', sessionId='sess_xxx', outputDir='output/', splitBy='page')

Template syntax (LINQ Reporting Engine, use 'ds' as data source prefix):
- Simple value: <<[ds.Name]>>
- Nested object: <<[ds.Customer.Address.City]>>
- Array iteration: <<foreach [item in ds.Items]>><<[item.Product]>>: <<[item.Price]>><</foreach>>")]
    public string Execute(
        [Description("Operation: create, create_from_template, convert, merge, split")]
        string operation,
        [Description("Session ID to read document from session (for convert, split, create_from_template)")]
        string? sessionId = null,
        [Description("Input file path (for convert, split)")]
        string? path = null,
        [Description("Output file path (for create, create_from_template, convert, merge)")]
        string? outputPath = null,
        [Description("Template file path (for create_from_template)")]
        string? templatePath = null,
        [Description("JSON data for template rendering (for create_from_template)")]
        string? dataJson = null,
        [Description("Output format: pdf, html, docx, txt, rtf, odt, epub, xps (for convert)")]
        string? format = null,
        [Description("Array of input file paths to merge (for merge)")]
        string[]? inputPaths = null,
        [Description(
            "Format mode when merging: KeepSourceFormatting, UseDestinationStyles, KeepDifferentStyles (default: KeepSourceFormatting)")]
        string importFormatMode = "KeepSourceFormatting",
        [Description("Unlink headers/footers after merge (default: false)")]
        bool unlinkHeadersFooters = false,
        [Description("Output directory for split files (for split)")]
        string? outputDir = null,
        [Description("Split by: section, page (default: section)")]
        string splitBy = "section",
        [Description("Initial content (for create)")]
        string? content = null,
        [Description("Create blank document (for create, default: false)")]
        bool skipInitialContent = false,
        [Description("Top margin in points (default: 70.87)")]
        double marginTop = 70.87,
        [Description("Bottom margin in points (default: 70.87)")]
        double marginBottom = 70.87,
        [Description("Left margin in points (default: 70.87)")]
        double marginLeft = 70.87,
        [Description("Right margin in points (default: 70.87)")]
        double marginRight = 70.87,
        [Description("Word compatibility mode: Word2019, Word2016, Word2013, Word2010, Word2007")]
        string compatibilityMode = "Word2019",
        [Description("Predefined paper size: A4, Letter, A3, Legal (default: A4)")]
        string paperSize = "A4",
        [Description("Page width in points (overrides paperSize)")]
        double? pageWidth = null,
        [Description("Page height in points (overrides paperSize)")]
        double? pageHeight = null,
        [Description("Header distance from page top in points (default: 35.4)")]
        double headerDistance = 35.4,
        [Description("Footer distance from page bottom in points (default: 35.4)")]
        double footerDistance = 35.4)
    {
        var parameters = BuildParameters(operation, sessionId, path, outputPath, templatePath, dataJson, format,
            inputPaths, importFormatMode, unlinkHeadersFooters, outputDir, splitBy, content, skipInitialContent,
            marginTop, marginBottom, marginLeft, marginRight, compatibilityMode, paperSize, pageWidth, pageHeight,
            headerDistance, footerDistance);

        var handler = _handlerRegistry.GetHandler(operation);

        // File operations don't use DocumentContext - create a minimal OperationContext
        var operationContext = new OperationContext<Document>
        {
            Document = null!,
            SessionManager = _sessionManager,
            IdentityAccessor = _identityAccessor,
            SessionId = sessionId,
            SourcePath = path,
            OutputPath = outputPath
        };

        return handler.Execute(operationContext, parameters);
    }

    /// <summary>
    ///     Builds OperationParameters from method parameters.
    /// </summary>
    private static OperationParameters BuildParameters(
        string operation,
        string? sessionId,
        string? path,
        string? outputPath,
        string? templatePath,
        string? dataJson,
        string? format,
        string[]? inputPaths,
        string importFormatMode,
        bool unlinkHeadersFooters,
        string? outputDir,
        string splitBy,
        string? content,
        bool skipInitialContent,
        double marginTop,
        double marginBottom,
        double marginLeft,
        double marginRight,
        string compatibilityMode,
        string paperSize,
        double? pageWidth,
        double? pageHeight,
        double headerDistance,
        double footerDistance)
    {
        var parameters = new OperationParameters();

        switch (operation.ToLower())
        {
            case "create":
                if (outputPath != null) parameters.Set("outputPath", outputPath);
                if (content != null) parameters.Set("content", content);
                parameters.Set("skipInitialContent", skipInitialContent);
                parameters.Set("marginTop", marginTop);
                parameters.Set("marginBottom", marginBottom);
                parameters.Set("marginLeft", marginLeft);
                parameters.Set("marginRight", marginRight);
                parameters.Set("compatibilityMode", compatibilityMode);
                parameters.Set("paperSize", paperSize);
                if (pageWidth.HasValue) parameters.Set("pageWidth", pageWidth.Value);
                if (pageHeight.HasValue) parameters.Set("pageHeight", pageHeight.Value);
                parameters.Set("headerDistance", headerDistance);
                parameters.Set("footerDistance", footerDistance);
                break;

            case "create_from_template":
                if (templatePath != null) parameters.Set("templatePath", templatePath);
                if (sessionId != null) parameters.Set("sessionId", sessionId);
                if (outputPath != null) parameters.Set("outputPath", outputPath);
                if (dataJson != null) parameters.Set("dataJson", dataJson);
                break;

            case "convert":
                if (path != null) parameters.Set("path", path);
                if (sessionId != null) parameters.Set("sessionId", sessionId);
                if (outputPath != null) parameters.Set("outputPath", outputPath);
                if (format != null) parameters.Set("format", format);
                break;

            case "merge":
                if (inputPaths != null) parameters.Set("inputPaths", inputPaths);
                if (outputPath != null) parameters.Set("outputPath", outputPath);
                parameters.Set("importFormatMode", importFormatMode);
                parameters.Set("unlinkHeadersFooters", unlinkHeadersFooters);
                break;

            case "split":
                if (path != null) parameters.Set("path", path);
                if (sessionId != null) parameters.Set("sessionId", sessionId);
                if (outputDir != null) parameters.Set("outputDir", outputDir);
                parameters.Set("splitBy", splitBy);
                break;
        }

        return parameters;
    }
}
