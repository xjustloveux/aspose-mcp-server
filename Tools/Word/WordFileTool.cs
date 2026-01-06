using System.ComponentModel;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Settings;
using AsposeMcpServer.Core.Helpers;
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
        return operation.ToLower() switch
        {
            "create" => CreateDocument(outputPath, content, skipInitialContent, marginTop, marginBottom, marginLeft,
                marginRight, compatibilityMode, paperSize, pageWidth, pageHeight, headerDistance, footerDistance),
            "create_from_template" => CreateFromTemplate(templatePath, sessionId, outputPath, dataJson),
            "convert" => ConvertDocument(path, sessionId, outputPath, format),
            "merge" => MergeDocuments(inputPaths, outputPath, importFormatMode, unlinkHeadersFooters),
            "split" => SplitDocument(path, sessionId, outputDir, splitBy),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Creates a new Word document with specified settings.
    /// </summary>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="content">Initial content for the document.</param>
    /// <param name="skipInitialContent">Whether to create a blank document.</param>
    /// <param name="marginTop">Top margin in points.</param>
    /// <param name="marginBottom">Bottom margin in points.</param>
    /// <param name="marginLeft">Left margin in points.</param>
    /// <param name="marginRight">Right margin in points.</param>
    /// <param name="compatibilityMode">Word compatibility mode.</param>
    /// <param name="paperSize">Predefined paper size.</param>
    /// <param name="pageWidth">Custom page width in points.</param>
    /// <param name="pageHeight">Custom page height in points.</param>
    /// <param name="headerDistance">Header distance from page top.</param>
    /// <param name="footerDistance">Footer distance from page bottom.</param>
    /// <returns>A message indicating the document was created successfully.</returns>
    /// <exception cref="ArgumentException">Thrown when outputPath is not provided.</exception>
    private static string CreateDocument(string? outputPath, string? content, bool skipInitialContent, double marginTop,
        double marginBottom, double marginLeft, double marginRight, string compatibilityMode, string paperSize,
        double? pageWidth, double? pageHeight, double headerDistance, double footerDistance)
    {
        if (string.IsNullOrEmpty(outputPath))
            throw new ArgumentException("outputPath is required for create operation");

        SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

        var outputDir = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(outputDir))
            Directory.CreateDirectory(outputDir);

        var doc = new Document();

        var wordVersion = compatibilityMode switch
        {
            "Word2019" => MsWordVersion.Word2019,
            "Word2016" => MsWordVersion.Word2016,
            "Word2013" => MsWordVersion.Word2013,
            "Word2010" => MsWordVersion.Word2010,
            "Word2007" => MsWordVersion.Word2007,
            _ => MsWordVersion.Word2019
        };
        doc.CompatibilityOptions.OptimizeFor(wordVersion);

        var section = doc.FirstSection;
        if (section != null)
        {
            var pageSetup = section.PageSetup;

            if (!string.IsNullOrEmpty(paperSize) && pageWidth == null && pageHeight == null)
            {
                pageSetup.PaperSize = paperSize.ToUpper() switch
                {
                    "A4" => PaperSize.A4,
                    "LETTER" => PaperSize.Letter,
                    "A3" => PaperSize.A3,
                    "LEGAL" => PaperSize.Legal,
                    _ => PaperSize.A4
                };
            }
            else if (pageWidth != null || pageHeight != null)
            {
                pageSetup.PaperSize = PaperSize.Custom;
                pageSetup.PageWidth = pageWidth ?? 595.3;
                pageSetup.PageHeight = pageHeight ?? 841.9;
            }
            else
            {
                pageSetup.PaperSize = PaperSize.A4;
            }

            pageSetup.TopMargin = marginTop;
            pageSetup.BottomMargin = marginBottom;
            pageSetup.LeftMargin = marginLeft;
            pageSetup.RightMargin = marginRight;
            pageSetup.HeaderDistance = headerDistance;
            pageSetup.FooterDistance = footerDistance;
        }

        var builder = new DocumentBuilder(doc);

        if (skipInitialContent)
        {
            if (doc.FirstSection is { Body: not null })
            {
                doc.FirstSection.Body.RemoveAllChildren();
                var firstPara = new Paragraph(doc)
                {
                    ParagraphFormat =
                    {
                        SpaceBefore = 0,
                        SpaceAfter = 0,
                        LineSpacing = 12
                    }
                };
                doc.FirstSection.Body.AppendChild(firstPara);
            }
        }
        else if (!string.IsNullOrEmpty(content))
        {
            builder.Write(content);
        }

        doc.Save(outputPath);
        return $"Word document created successfully at: {outputPath}";
    }

    /// <summary>
    ///     Creates a document from a template using LINQ Reporting Engine.
    /// </summary>
    /// <param name="templatePath">The template file path.</param>
    /// <param name="sessionId">The session ID for reading template from session.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="dataJson">The JSON data for template rendering.</param>
    /// <returns>A message indicating the document was created successfully.</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when neither templatePath nor sessionId is provided, outputPath is not provided, or dataJson is not
    ///     provided.
    /// </exception>
    /// <exception cref="InvalidOperationException">Thrown when session management is not enabled or document cloning fails.</exception>
    /// <exception cref="FileNotFoundException">Thrown when the template file is not found.</exception>
    private string CreateFromTemplate(string? templatePath, string? sessionId, string? outputPath, string? dataJson)
    {
        if (string.IsNullOrEmpty(templatePath) && string.IsNullOrEmpty(sessionId))
            throw new ArgumentException(
                "Either templatePath or sessionId is required for create_from_template operation");
        if (string.IsNullOrEmpty(outputPath))
            throw new ArgumentException("outputPath is required for create_from_template operation");

        SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

        var outputDir = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(outputDir))
            Directory.CreateDirectory(outputDir);

        if (string.IsNullOrEmpty(dataJson))
            throw new ArgumentException("dataJson parameter is required for create_from_template");

        Document doc;
        string templateSource;

        if (!string.IsNullOrEmpty(sessionId))
        {
            if (_sessionManager == null)
                throw new InvalidOperationException("Session management is not enabled");

            var identity = _identityAccessor?.GetCurrentIdentity() ?? SessionIdentity.GetAnonymous();
            var sessionDoc = _sessionManager.GetDocument<Document>(sessionId, identity);
            doc = sessionDoc.Clone() ?? throw new InvalidOperationException("Failed to clone document from session");
            templateSource = $"session {sessionId}";
        }
        else
        {
            SecurityHelper.ValidateFilePath(templatePath!, "templatePath", true);
            if (!File.Exists(templatePath))
                throw new FileNotFoundException($"Template file not found: {templatePath}");
            doc = new Document(templatePath);
            templateSource = Path.GetFileName(templatePath);
        }

        var engine = new ReportingEngine
        {
            Options = ReportBuildOptions.AllowMissingMembers | ReportBuildOptions.RemoveEmptyParagraphs
        };

        using var jsonStream = new MemoryStream(Encoding.UTF8.GetBytes(dataJson));
        var loadOptions = new JsonDataLoadOptions
        {
            ExactDateTimeParseFormats = ["yyyy-MM-dd", "yyyy-MM-ddTHH:mm:ss"],
            SimpleValueParseMode = JsonSimpleValueParseMode.Strict
        };
        var dataSource = new JsonDataSource(jsonStream, loadOptions);

        engine.BuildReport(doc, dataSource, "ds");

        doc.Save(outputPath);
        return $"Document created from template ({templateSource}) using LINQ Reporting Engine: {outputPath}";
    }

    /// <summary>
    ///     Converts a Word document to another format.
    /// </summary>
    /// <param name="path">The input file path.</param>
    /// <param name="sessionId">The session ID for reading document from session.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="format">The target format (pdf, html, docx, txt, rtf, odt, epub, xps).</param>
    /// <returns>A message indicating the conversion was successful.</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when neither path nor sessionId is provided, outputPath is not provided, or format is unsupported.
    /// </exception>
    /// <exception cref="InvalidOperationException">Thrown when session management is not enabled.</exception>
    private string ConvertDocument(string? path, string? sessionId, string? outputPath, string? format)
    {
        if (string.IsNullOrEmpty(path) && string.IsNullOrEmpty(sessionId))
            throw new ArgumentException("Either path or sessionId is required for convert operation");
        if (string.IsNullOrEmpty(outputPath))
            throw new ArgumentException("outputPath is required for convert operation");

        SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

        var outputDir = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(outputDir))
            Directory.CreateDirectory(outputDir);

        var formatLower = format?.ToLower();
        if (string.IsNullOrEmpty(formatLower))
        {
            var extension = Path.GetExtension(outputPath).TrimStart('.').ToLower();
            formatLower = extension switch
            {
                "pdf" => "pdf",
                "html" or "htm" => "html",
                "docx" => "docx",
                "doc" => "doc",
                "txt" => "txt",
                "rtf" => "rtf",
                "odt" => "odt",
                "epub" => "epub",
                "xps" => "xps",
                _ => throw new ArgumentException(
                    $"Cannot infer format from extension '.{extension}'. Please specify format parameter.")
            };
        }

        Document doc;
        string sourceDescription;

        if (!string.IsNullOrEmpty(sessionId))
        {
            if (_sessionManager == null)
                throw new InvalidOperationException("Session management is not enabled");

            var identity = _identityAccessor?.GetCurrentIdentity() ?? SessionIdentity.GetAnonymous();
            doc = _sessionManager.GetDocument<Document>(sessionId, identity);
            sourceDescription = $"session {sessionId}";
        }
        else
        {
            SecurityHelper.ValidateFilePath(path!, allowAbsolutePaths: true);
            doc = new Document(path);
            sourceDescription = path!;
        }

        var saveFormat = formatLower switch
        {
            "pdf" => SaveFormat.Pdf,
            "html" => SaveFormat.Html,
            "docx" => SaveFormat.Docx,
            "doc" => SaveFormat.Doc,
            "txt" => SaveFormat.Text,
            "rtf" => SaveFormat.Rtf,
            "odt" => SaveFormat.Odt,
            "epub" => SaveFormat.Epub,
            "xps" => SaveFormat.Xps,
            _ => throw new ArgumentException($"Unsupported format: {format}")
        };

        doc.Save(outputPath, saveFormat);
        return $"Document converted from {sourceDescription} to {outputPath} ({formatLower})";
    }

    /// <summary>
    ///     Merges multiple Word documents into one.
    /// </summary>
    /// <param name="inputPaths">Array of input file paths to merge.</param>
    /// <param name="outputPath">The output file path for the merged document.</param>
    /// <param name="importFormatModeStr">Format mode: KeepSourceFormatting, UseDestinationStyles, KeepDifferentStyles.</param>
    /// <param name="unlinkHeadersFooters">Whether to unlink headers/footers after merge.</param>
    /// <returns>A message indicating the merge was successful with document count.</returns>
    /// <exception cref="ArgumentException">Thrown when inputPaths is empty or outputPath is not provided.</exception>
    private static string MergeDocuments(string[]? inputPaths, string? outputPath, string importFormatModeStr,
        bool unlinkHeadersFooters)
    {
        if (inputPaths == null || inputPaths.Length == 0)
            throw new ArgumentException("inputPaths is required for merge operation");
        if (string.IsNullOrEmpty(outputPath))
            throw new ArgumentException("outputPath is required for merge operation");

        SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

        var outputDir = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(outputDir))
            Directory.CreateDirectory(outputDir);

        foreach (var inputPath in inputPaths)
            SecurityHelper.ValidateFilePath(inputPath, "inputPaths", true);

        var importFormatMode = importFormatModeStr switch
        {
            "UseDestinationStyles" => ImportFormatMode.UseDestinationStyles,
            "KeepDifferentStyles" => ImportFormatMode.KeepDifferentStyles,
            _ => ImportFormatMode.KeepSourceFormatting
        };

        var mergedDoc = new Document(inputPaths[0]);

        for (var i = 1; i < inputPaths.Length; i++)
        {
            var doc = new Document(inputPaths[i]);
            mergedDoc.AppendDocument(doc, importFormatMode);
        }

        if (unlinkHeadersFooters)
            foreach (var section in mergedDoc.Sections.Cast<Section>())
                section.HeadersFooters.LinkToPrevious(false);

        mergedDoc.Save(outputPath);
        return $"Merged {inputPaths.Length} documents into: {outputPath} (format mode: {importFormatModeStr})";
    }

    /// <summary>
    ///     Splits a Word document by sections or pages.
    /// </summary>
    /// <param name="path">The input file path.</param>
    /// <param name="sessionId">The session ID for reading document from session.</param>
    /// <param name="outputDir">The output directory for split files.</param>
    /// <param name="splitBy">Split method: section or page.</param>
    /// <returns>A message indicating the split was successful with file count.</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when neither path nor sessionId is provided, or outputDir is not provided.
    /// </exception>
    /// <exception cref="InvalidOperationException">Thrown when session management is not enabled.</exception>
    private string SplitDocument(string? path, string? sessionId, string? outputDir, string splitBy)
    {
        if (string.IsNullOrEmpty(path) && string.IsNullOrEmpty(sessionId))
            throw new ArgumentException("Either path or sessionId is required for split operation");
        if (string.IsNullOrEmpty(outputDir))
            throw new ArgumentException("outputDir is required for split operation");

        SecurityHelper.ValidateFilePath(outputDir, "outputDir", true);
        Directory.CreateDirectory(outputDir);

        Document doc;
        string fileBaseName;

        if (!string.IsNullOrEmpty(sessionId))
        {
            if (_sessionManager == null)
                throw new InvalidOperationException("Session management is not enabled");

            var identity = _identityAccessor?.GetCurrentIdentity() ?? SessionIdentity.GetAnonymous();
            doc = _sessionManager.GetDocument<Document>(sessionId, identity);
            fileBaseName = $"session_{sessionId}";
        }
        else
        {
            SecurityHelper.ValidateFilePath(path!, allowAbsolutePaths: true);
            doc = new Document(path);
            fileBaseName = SecurityHelper.SanitizeFileName(Path.GetFileNameWithoutExtension(path!));
        }

        if (splitBy.ToLower() == "section")
        {
            for (var i = 0; i < doc.Sections.Count; i++)
            {
                var sectionDoc = new Document();
                sectionDoc.RemoveAllChildren();
                sectionDoc.AppendChild(sectionDoc.ImportNode(doc.Sections[i], true));

                var output = Path.Combine(outputDir, $"{fileBaseName}_section_{i + 1}.docx");
                sectionDoc.Save(output);
            }

            return $"Document split into {doc.Sections.Count} sections in: {outputDir}";
        }

        doc.UpdatePageLayout();

        var pageCount = doc.PageCount;
        for (var i = 0; i < pageCount; i++)
        {
            var pageDoc = doc.ExtractPages(i, 1);
            var output = Path.Combine(outputDir, $"{fileBaseName}_page_{i + 1}.docx");
            pageDoc.Save(output);
        }

        return $"Document split into {pageCount} pages in: {outputDir}";
    }
}