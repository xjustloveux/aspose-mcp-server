using System.ComponentModel;
using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Layout;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Unified tool for page operations in Word documents
///     Merges: WordSetPageMarginsTool, WordSetPageOrientationTool, WordSetPageSizeTool,
///     WordSetPageNumberTool, WordSetPageSetupTool, WordDeletePageTool, WordInsertBlankPageTool, WordAddPageBreakTool
/// </summary>
[McpServerToolType]
public class WordPageTool
{
    /// <summary>
    ///     Identity accessor for session isolation
    /// </summary>
    private readonly ISessionIdentityAccessor? _identityAccessor;

    /// <summary>
    ///     Session manager for document session operations
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the WordPageTool class
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document operations</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation</param>
    public WordPageTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
    }

    /// <summary>
    ///     Executes a Word page operation (set_margins, set_orientation, set_size, set_page_number, set_page_setup,
    ///     delete_page, insert_blank_page, add_page_break).
    /// </summary>
    /// <param name="operation">
    ///     The operation to perform: set_margins, set_orientation, set_size, set_page_number,
    ///     set_page_setup, delete_page, insert_blank_page, add_page_break.
    /// </param>
    /// <param name="path">Word document file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="top">Top margin in points (72 pts = 1 inch).</param>
    /// <param name="bottom">Bottom margin in points.</param>
    /// <param name="left">Left margin in points.</param>
    /// <param name="right">Right margin in points.</param>
    /// <param name="orientation">Page orientation: Portrait or Landscape.</param>
    /// <param name="width">Page width in points (72 pts = 1 inch).</param>
    /// <param name="height">Page height in points.</param>
    /// <param name="paperSize">Predefined paper size: A4, Letter, Legal, A3, A5.</param>
    /// <param name="pageNumberFormat">Page number format: arabic, roman, letter.</param>
    /// <param name="startingPageNumber">Starting page number.</param>
    /// <param name="sectionIndex">Section index (0-based).</param>
    /// <param name="sectionIndices">Array of section indices.</param>
    /// <param name="pageIndex">Page index to delete (0-based).</param>
    /// <param name="insertAtPageIndex">Page index to insert blank page at (0-based).</param>
    /// <param name="paragraphIndex">Paragraph index to insert page break after (0-based).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "word_page")]
    [Description(
        @"Manage page settings in Word documents. Supports 8 operations: set_margins, set_orientation, set_size, set_page_number, set_page_setup, delete_page, insert_blank_page, add_page_break.

Usage examples:
- Set margins: word_page(operation='set_margins', path='doc.docx', top=72, bottom=72, left=72, right=72)
- Set orientation: word_page(operation='set_orientation', path='doc.docx', orientation='landscape')
- Set page size: word_page(operation='set_size', path='doc.docx', width=792, height=612)
- Set page number: word_page(operation='set_page_number', path='doc.docx', startingPageNumber=1)
- Delete page: word_page(operation='delete_page', path='doc.docx', pageIndex=1)
- Insert blank page: word_page(operation='insert_blank_page', path='doc.docx', insertAtPageIndex=2)
- Add page break: word_page(operation='add_page_break', path='doc.docx', paragraphIndex=10)")]
    public string Execute(
        [Description(
            "Operation: set_margins, set_orientation, set_size, set_page_number, set_page_setup, delete_page, insert_blank_page, add_page_break")]
        string operation,
        [Description("Document file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Top margin in points (72 pts = 1 inch)")]
        double? top = null,
        [Description("Bottom margin in points")]
        double? bottom = null,
        [Description("Left margin in points")] double? left = null,
        [Description("Right margin in points")]
        double? right = null,
        [Description("Page orientation: Portrait or Landscape")]
        string? orientation = null,
        [Description("Page width in points (72 pts = 1 inch)")]
        double? width = null,
        [Description("Page height in points")] double? height = null,
        [Description("Predefined paper size: A4, Letter, Legal, A3, A5")]
        string? paperSize = null,
        [Description("Page number format: arabic, roman, letter")]
        string? pageNumberFormat = null,
        [Description("Starting page number")] int? startingPageNumber = null,
        [Description("Section index (0-based)")]
        int? sectionIndex = null,
        [Description("Array of section indices (overrides sectionIndex)")]
        JsonArray? sectionIndices = null,
        [Description("Page index to delete (0-based)")]
        int? pageIndex = null,
        [Description("Page index to insert blank page at (0-based)")]
        int? insertAtPageIndex = null,
        [Description("Paragraph index to insert page break after (0-based)")]
        int? paragraphIndex = null)
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);

        return operation.ToLower() switch
        {
            "set_margins" => SetMargins(ctx, outputPath, top, bottom, left, right, sectionIndex, sectionIndices),
            "set_orientation" => SetOrientation(ctx, outputPath, orientation, sectionIndex, sectionIndices),
            "set_size" => SetSize(ctx, outputPath, width, height, paperSize, sectionIndex, sectionIndices),
            "set_page_number" => SetPageNumber(ctx, outputPath, pageNumberFormat, startingPageNumber, sectionIndex),
            "set_page_setup" => SetPageSetup(ctx, outputPath, top, bottom, left, right, orientation, sectionIndex),
            "delete_page" => DeletePage(ctx, outputPath, pageIndex),
            "insert_blank_page" => InsertBlankPage(ctx, outputPath, insertAtPageIndex),
            "add_page_break" => AddPageBreak(ctx, outputPath, paragraphIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Gets the list of section indices to operate on based on provided parameters.
    /// </summary>
    /// <param name="doc">The Word document.</param>
    /// <param name="sectionIndex">Optional single section index.</param>
    /// <param name="sectionIndices">Optional array of section indices.</param>
    /// <param name="validateRange">Whether to validate that indices are within range.</param>
    /// <returns>A list of section indices to operate on.</returns>
    /// <exception cref="ArgumentException">Thrown when section index is out of range.</exception>
    private static List<int> GetTargetSections(Document doc, int? sectionIndex, JsonArray? sectionIndices,
        bool validateRange = true)
    {
        if (sectionIndices is { Count: > 0 })
        {
            var indices = sectionIndices
                .Select(s => s?.GetValue<int>())
                .Where(s => s.HasValue)
                .Select(s => s!.Value)
                .ToList();

            if (validateRange)
                foreach (var idx in indices)
                    if (idx < 0 || idx >= doc.Sections.Count)
                        throw new ArgumentException(
                            $"sectionIndex {idx} must be between 0 and {doc.Sections.Count - 1}");

            return indices;
        }

        if (sectionIndex.HasValue)
        {
            if (validateRange && (sectionIndex.Value < 0 || sectionIndex.Value >= doc.Sections.Count))
                throw new ArgumentException($"sectionIndex must be between 0 and {doc.Sections.Count - 1}");
            return [sectionIndex.Value];
        }

        return Enumerable.Range(0, doc.Sections.Count).ToList();
    }

    /// <summary>
    ///     Sets page margins for the specified sections.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="top">The top margin in points.</param>
    /// <param name="bottom">The bottom margin in points.</param>
    /// <param name="left">The left margin in points.</param>
    /// <param name="right">The right margin in points.</param>
    /// <param name="sectionIndex">Optional section index.</param>
    /// <param name="sectionIndices">Optional array of section indices.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    private static string SetMargins(DocumentContext<Document> ctx, string? outputPath, double? top, double? bottom,
        double? left, double? right, int? sectionIndex, JsonArray? sectionIndices)
    {
        var doc = ctx.Document;
        var sectionsToUpdate = GetTargetSections(doc, sectionIndex, sectionIndices);

        foreach (var idx in sectionsToUpdate)
        {
            var pageSetup = doc.Sections[idx].PageSetup;
            if (top.HasValue) pageSetup.TopMargin = top.Value;
            if (bottom.HasValue) pageSetup.BottomMargin = bottom.Value;
            if (left.HasValue) pageSetup.LeftMargin = left.Value;
            if (right.HasValue) pageSetup.RightMargin = right.Value;
        }

        ctx.Save(outputPath);
        return $"Page margins updated for {sectionsToUpdate.Count} section(s)\n{ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Sets page orientation (portrait or landscape) for the specified sections.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="orientation">The page orientation (Portrait or Landscape).</param>
    /// <param name="sectionIndex">Optional section index.</param>
    /// <param name="sectionIndices">Optional array of section indices.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when orientation is not specified.</exception>
    private static string SetOrientation(DocumentContext<Document> ctx, string? outputPath, string? orientation,
        int? sectionIndex, JsonArray? sectionIndices)
    {
        if (string.IsNullOrEmpty(orientation))
            throw new ArgumentException("orientation parameter is required for set_orientation operation");

        var doc = ctx.Document;
        var orientationEnum = orientation.ToLower() == "landscape" ? Orientation.Landscape : Orientation.Portrait;
        var sectionsToUpdate = GetTargetSections(doc, sectionIndex, sectionIndices);

        foreach (var idx in sectionsToUpdate)
            doc.Sections[idx].PageSetup.Orientation = orientationEnum;

        ctx.Save(outputPath);
        return
            $"Page orientation set to {orientation} for {sectionsToUpdate.Count} section(s)\n{ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Sets page size using custom dimensions or predefined paper sizes.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="width">The page width in points.</param>
    /// <param name="height">The page height in points.</param>
    /// <param name="paperSize">The predefined paper size (A4, Letter, Legal, A3, A5).</param>
    /// <param name="sectionIndex">Optional section index.</param>
    /// <param name="sectionIndices">Optional array of section indices.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when neither paperSize nor width/height are provided.</exception>
    private static string SetSize(DocumentContext<Document> ctx, string? outputPath, double? width, double? height,
        string? paperSize, int? sectionIndex, JsonArray? sectionIndices)
    {
        var doc = ctx.Document;
        var sectionsToUpdate = GetTargetSections(doc, sectionIndex, sectionIndices);

        foreach (var idx in sectionsToUpdate)
        {
            var pageSetup = doc.Sections[idx].PageSetup;

            if (!string.IsNullOrEmpty(paperSize))
            {
                pageSetup.PaperSize = paperSize.ToUpper() switch
                {
                    "A4" => PaperSize.A4,
                    "LETTER" => PaperSize.Letter,
                    "LEGAL" => PaperSize.Legal,
                    "A3" => PaperSize.A3,
                    "A5" => PaperSize.A5,
                    _ => PaperSize.A4
                };
            }
            else if (width.HasValue && height.HasValue)
            {
                pageSetup.PageWidth = width.Value;
                pageSetup.PageHeight = height.Value;
            }
            else
            {
                throw new ArgumentException("Either paperSize or both width and height must be provided");
            }
        }

        ctx.Save(outputPath);
        return $"Page size updated for {sectionsToUpdate.Count} section(s)\n{ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Sets page number format and starting number for a section.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="pageNumberFormat">The page number format (arabic, roman, letter).</param>
    /// <param name="startingPageNumber">The starting page number.</param>
    /// <param name="sectionIndex">Optional section index.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when section index is out of range.</exception>
    private static string SetPageNumber(DocumentContext<Document> ctx, string? outputPath, string? pageNumberFormat,
        int? startingPageNumber, int? sectionIndex)
    {
        var doc = ctx.Document;
        List<int> sectionsToUpdate;

        if (sectionIndex.HasValue)
        {
            if (sectionIndex.Value < 0 || sectionIndex.Value >= doc.Sections.Count)
                throw new ArgumentException($"sectionIndex must be between 0 and {doc.Sections.Count - 1}");
            sectionsToUpdate = [sectionIndex.Value];
        }
        else
        {
            sectionsToUpdate = [0];
        }

        foreach (var idx in sectionsToUpdate)
        {
            var pageSetup = doc.Sections[idx].PageSetup;

            if (!string.IsNullOrEmpty(pageNumberFormat))
            {
                var numStyle = pageNumberFormat.ToLower() switch
                {
                    "roman" => NumberStyle.UppercaseRoman,
                    "letter" => NumberStyle.UppercaseLetter,
                    _ => NumberStyle.Arabic
                };
                pageSetup.PageNumberStyle = numStyle;
            }

            if (startingPageNumber.HasValue)
            {
                pageSetup.RestartPageNumbering = true;
                pageSetup.PageStartingNumber = startingPageNumber.Value;
            }
        }

        ctx.Save(outputPath);
        return
            $"Page number settings updated for {sectionsToUpdate.Count} section(s)\n{ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Sets multiple page setup options (margins and orientation) for a section.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="top">The top margin in points.</param>
    /// <param name="bottom">The bottom margin in points.</param>
    /// <param name="left">The left margin in points.</param>
    /// <param name="right">The right margin in points.</param>
    /// <param name="orientation">The page orientation.</param>
    /// <param name="sectionIndex">Optional section index.</param>
    /// <returns>A message indicating the changes made.</returns>
    /// <exception cref="ArgumentException">Thrown when section index is out of range.</exception>
    private static string SetPageSetup(DocumentContext<Document> ctx, string? outputPath, double? top, double? bottom,
        double? left, double? right, string? orientation, int? sectionIndex)
    {
        var doc = ctx.Document;
        var idx = sectionIndex ?? 0;

        if (idx < 0 || idx >= doc.Sections.Count)
            throw new ArgumentException($"sectionIndex must be between 0 and {doc.Sections.Count - 1}");

        var pageSetup = doc.Sections[idx].PageSetup;
        List<string> changes = [];

        if (top.HasValue)
        {
            pageSetup.TopMargin = top.Value;
            changes.Add($"Top margin: {top.Value}");
        }

        if (bottom.HasValue)
        {
            pageSetup.BottomMargin = bottom.Value;
            changes.Add($"Bottom margin: {bottom.Value}");
        }

        if (left.HasValue)
        {
            pageSetup.LeftMargin = left.Value;
            changes.Add($"Left margin: {left.Value}");
        }

        if (right.HasValue)
        {
            pageSetup.RightMargin = right.Value;
            changes.Add($"Right margin: {right.Value}");
        }

        if (!string.IsNullOrEmpty(orientation))
        {
            pageSetup.Orientation = orientation.ToLower() == "landscape" ? Orientation.Landscape : Orientation.Portrait;
            changes.Add($"Orientation: {orientation}");
        }

        ctx.Save(outputPath);
        return $"Page setup updated: {string.Join(", ", changes)}";
    }

    /// <summary>
    ///     Deletes a specific page from the document by extracting and recombining pages.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="pageIndex">The page index to delete (0-based).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when page index is not provided or out of range.</exception>
    /// <exception cref="InvalidOperationException">Thrown when output path is required but not provided.</exception>
    private static string DeletePage(DocumentContext<Document> ctx, string? outputPath, int? pageIndex)
    {
        if (!pageIndex.HasValue)
            throw new ArgumentException("pageIndex parameter is required for delete_page operation");

        var doc = ctx.Document;
        var pageCount = doc.PageCount;

        if (pageIndex.Value < 0 || pageIndex.Value >= pageCount)
            throw new ArgumentException(
                $"pageIndex must be between 0 and {pageCount - 1} (document has {pageCount} pages)");

        var resultDoc = new Document();
        resultDoc.RemoveAllChildren();

        if (pageIndex.Value > 0)
        {
            var beforePages = doc.ExtractPages(0, pageIndex.Value);
            foreach (var section in beforePages.Sections.Cast<Section>())
                resultDoc.AppendChild(resultDoc.ImportNode(section, true));
        }

        if (pageIndex.Value < pageCount - 1)
        {
            var afterPages = doc.ExtractPages(pageIndex.Value + 1, pageCount - pageIndex.Value - 1);
            foreach (var section in afterPages.Sections.Cast<Section>())
                resultDoc.AppendChild(resultDoc.ImportNode(section, true));
        }

        // For file mode, save the result document directly
        if (!ctx.IsSession)
        {
            var savePath = outputPath ?? throw new InvalidOperationException("Output path required for file mode");
            resultDoc.Save(savePath);
            return
                $"Page {pageIndex.Value} deleted successfully (document now has {resultDoc.PageCount} pages)\nOutput: {savePath}";
        }

        // For session mode, we need to update the session document
        // Clear the current document and copy content from result
        doc.RemoveAllChildren();
        foreach (var section in resultDoc.Sections.Cast<Section>())
            doc.AppendChild(doc.ImportNode(section, true));

        ctx.Save(outputPath);
        return
            $"Page {pageIndex.Value} deleted successfully (document now has {doc.PageCount} pages)\n{ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Inserts a blank page at the specified position in the document.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="insertAtPageIndex">The page index to insert at (0-based).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when insert page index is out of range.</exception>
    private static string InsertBlankPage(DocumentContext<Document> ctx, string? outputPath, int? insertAtPageIndex)
    {
        var doc = ctx.Document;
        var builder = new DocumentBuilder(doc);

        if (insertAtPageIndex is > 0)
        {
            var pageCount = doc.PageCount;
            if (insertAtPageIndex.Value > pageCount)
                throw new ArgumentException(
                    $"insertAtPageIndex must be between 0 and {pageCount} (document has {pageCount} pages)");

            var layoutCollector = new LayoutCollector(doc);
            var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();

            Paragraph? targetParagraph = null;
            foreach (var para in paragraphs)
            {
                var paraPage = layoutCollector.GetStartPageIndex(para);
                if (paraPage == insertAtPageIndex.Value + 1)
                {
                    targetParagraph = para;
                    break;
                }
            }

            if (targetParagraph != null)
            {
                builder.MoveTo(targetParagraph);
                builder.InsertBreak(BreakType.PageBreak);
            }
            else
            {
                builder.MoveToDocumentEnd();
                builder.InsertBreak(BreakType.PageBreak);
            }
        }
        else
        {
            builder.MoveToDocumentEnd();
            builder.InsertBreak(BreakType.PageBreak);
        }

        ctx.Save(outputPath);
        return $"Blank page inserted at page {insertAtPageIndex ?? doc.PageCount}\n{ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Adds a page break at the specified paragraph or at document end.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="paragraphIndex">The paragraph index to insert page break after (0-based).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when paragraph index is out of range.</exception>
    private static string AddPageBreak(DocumentContext<Document> ctx, string? outputPath, int? paragraphIndex)
    {
        var doc = ctx.Document;
        var builder = new DocumentBuilder(doc);

        if (paragraphIndex.HasValue)
        {
            var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
            if (paragraphIndex.Value < 0 || paragraphIndex.Value >= paragraphs.Count)
                throw new ArgumentException($"paragraphIndex must be between 0 and {paragraphs.Count - 1}");

            builder.MoveTo(paragraphs[paragraphIndex.Value]);
            builder.InsertBreak(BreakType.PageBreak);
        }
        else
        {
            builder.MoveToDocumentEnd();
            builder.InsertBreak(BreakType.PageBreak);
        }

        ctx.Save(outputPath);
        var location = paragraphIndex.HasValue ? $"after paragraph {paragraphIndex.Value}" : "at document end";
        return $"Page break added {location}\n{ctx.GetOutputMessage(outputPath)}";
    }
}