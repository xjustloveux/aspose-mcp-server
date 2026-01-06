using System.ComponentModel;
using System.Text.Json;
using System.Text.RegularExpressions;
using Aspose.Pdf;
using Aspose.Pdf.Text;
using AsposeMcpServer.Core.Helpers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Pdf;

/// <summary>
///     Tool for managing text in PDF documents (add, edit, extract)
/// </summary>
[McpServerToolType]
public class PdfTextTool
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
    ///     Initializes a new instance of the <see cref="PdfTextTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional session identity accessor for session isolation.</param>
    public PdfTextTool(DocumentSessionManager? sessionManager = null, ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
    }

    /// <summary>
    ///     Executes a PDF text operation (add, edit, extract).
    /// </summary>
    /// <param name="operation">The operation to perform: add, edit, extract.</param>
    /// <param name="path">PDF file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="pageIndex">Page index (1-based).</param>
    /// <param name="text">Text to add (required for add).</param>
    /// <param name="x">X position in PDF coordinates (for add).</param>
    /// <param name="y">Y position in PDF coordinates (for add).</param>
    /// <param name="fontName">Font name (for add).</param>
    /// <param name="fontSize">Font size (for add).</param>
    /// <param name="oldText">Text to replace (required for edit).</param>
    /// <param name="newText">New text (required for edit).</param>
    /// <param name="replaceAll">Replace all occurrences (for edit).</param>
    /// <param name="includeFontInfo">Include font information (for extract).</param>
    /// <param name="extractionMode">Text extraction mode (for extract).</param>
    /// <returns>A message indicating the result of the operation, or JSON data for extract operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "pdf_text")]
    [Description(@"Manage text in PDF documents. Supports 3 operations: add, edit, extract.

Usage examples:
- Add text: pdf_text(operation='add', path='doc.pdf', pageIndex=1, text='Hello World', x=100, y=700)
- Add text with font: pdf_text(operation='add', path='doc.pdf', pageIndex=1, text='Hello', x=100, y=700, fontName='Arial', fontSize=14)
- Edit text: pdf_text(operation='edit', path='doc.pdf', pageIndex=1, oldText='old', newText='new')
- Edit all occurrences: pdf_text(operation='edit', path='doc.pdf', pageIndex=1, oldText='old', newText='new', replaceAll=true)
- Extract text: pdf_text(operation='extract', path='doc.pdf', pageIndex=1)
- Extract with font info: pdf_text(operation='extract', path='doc.pdf', pageIndex=1, includeFontInfo=true)
- Extract raw text: pdf_text(operation='extract', path='doc.pdf', pageIndex=1, extractionMode='raw')")]
    public string Execute(
        [Description("Operation: add, edit, extract")]
        string operation,
        [Description("PDF file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Page index (1-based)")] int pageIndex = 1,
        [Description("Text to add (required for add)")]
        string? text = null,
        [Description("X position in PDF coordinates (for add, default: 100)")]
        double x = 100,
        [Description("Y position in PDF coordinates (for add, default: 700)")]
        double y = 700,
        [Description("Font name (for add, default: 'Arial')")]
        string fontName = "Arial",
        [Description("Font size (for add, default: 12)")]
        double fontSize = 12,
        [Description("Text to replace (required for edit)")]
        string? oldText = null,
        [Description("New text (required for edit)")]
        string? newText = null,
        [Description("Replace all occurrences (for edit, default: false)")]
        bool replaceAll = false,
        [Description("Include font information (for extract, default: false)")]
        bool includeFontInfo = false,
        [Description("Text extraction mode (for extract, default: 'pure')")]
        string extractionMode = "pure")
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);

        return operation.ToLower() switch
        {
            "add" => AddText(ctx, outputPath, pageIndex, text, x, y, fontName, fontSize),
            "edit" => EditText(ctx, outputPath, pageIndex, oldText, newText, replaceAll),
            "extract" => ExtractText(ctx, pageIndex, includeFontInfo, extractionMode),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds text to the specified page of the PDF document.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="pageIndex">The 1-based page index.</param>
    /// <param name="text">The text to add.</param>
    /// <param name="x">The X position in PDF coordinates.</param>
    /// <param name="y">The Y position in PDF coordinates.</param>
    /// <param name="fontName">The font name to use.</param>
    /// <param name="fontSize">The font size in points.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when text is empty or page index is invalid.</exception>
    private static string AddText(DocumentContext<Document> ctx, string? outputPath, int pageIndex, string? text,
        double x, double y, string fontName, double fontSize)
    {
        if (string.IsNullOrEmpty(text))
            throw new ArgumentException("text is required for add operation");

        var document = ctx.Document;
        if (pageIndex < 1 || pageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var page = document.Pages[pageIndex];
        var textFragment = new TextFragment(text)
        {
            Position = new Position(x, y)
        };

        FontHelper.Pdf.ApplyFontSettings(textFragment.TextState, fontName, fontSize);

        var textBuilder = new TextBuilder(page);
        textBuilder.AppendText(textFragment);

        ctx.Save(outputPath);
        return $"Added text to page {pageIndex}. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Edits (replaces) text on the specified page of the PDF document.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="pageIndex">The 1-based page index.</param>
    /// <param name="oldText">The text to find and replace.</param>
    /// <param name="newText">The replacement text.</param>
    /// <param name="replaceAll">Whether to replace all occurrences.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when text is not found or page index is invalid.</exception>
    private static string EditText(DocumentContext<Document> ctx, string? outputPath, int pageIndex, string? oldText,
        string? newText, bool replaceAll)
    {
        if (string.IsNullOrEmpty(oldText))
            throw new ArgumentException("oldText is required for edit operation");
        if (string.IsNullOrEmpty(newText))
            throw new ArgumentException("newText is required for edit operation");

        var document = ctx.Document;
        if (pageIndex < 1 || pageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var page = document.Pages[pageIndex];

        var textFragmentAbsorber = new TextFragmentAbsorber(oldText);
        page.Accept(textFragmentAbsorber);

        var fragments = textFragmentAbsorber.TextFragments;
        var normalizedOldText = Regex.Replace(oldText, @"\s+", " ").Trim();

        if (fragments.Count == 0 && normalizedOldText != oldText)
        {
            textFragmentAbsorber = new TextFragmentAbsorber(normalizedOldText);
            page.Accept(textFragmentAbsorber);
            fragments = textFragmentAbsorber.TextFragments;
        }

        if (fragments.Count == 0)
        {
            var textAbsorber = new TextAbsorber();
            page.Accept(textAbsorber);
            var pageText = textAbsorber.Text ?? "";
            var preview = pageText.Length > 200 ? pageText[..200] + "..." : pageText;
            throw new ArgumentException(
                $"Text '{oldText}' not found on page {pageIndex}. Page text preview: {preview}");
        }

        var finalReplaceCount = replaceAll ? fragments.Count : 1;
        var replacedCount = 0;

        foreach (var fragment in fragments)
        {
            if (replacedCount >= finalReplaceCount)
                break;
            fragment.Text = newText;
            replacedCount++;
        }

        ctx.Save(outputPath);
        return
            $"Replaced {replacedCount} occurrence(s) of '{oldText}' on page {pageIndex}. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Extracts text from the specified page of the PDF document.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="pageIndex">The 1-based page index.</param>
    /// <param name="includeFontInfo">Whether to include font information in the output.</param>
    /// <param name="extractionMode">The text extraction mode (pure or raw).</param>
    /// <returns>A JSON string containing the extracted text.</returns>
    /// <exception cref="ArgumentException">Thrown when the page index is invalid.</exception>
    private static string ExtractText(DocumentContext<Document> ctx, int pageIndex, bool includeFontInfo,
        string extractionMode)
    {
        var document = ctx.Document;
        if (pageIndex < 1 || pageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var page = document.Pages[pageIndex];

        var textAbsorber = new TextAbsorber();
        if (extractionMode.ToLower() == "raw")
            textAbsorber.ExtractionOptions = new TextExtractionOptions(TextExtractionOptions.TextFormattingMode.Raw);

        page.Accept(textAbsorber);

        if (includeFontInfo)
        {
            var textFragmentAbsorber = new TextFragmentAbsorber();
            page.Accept(textFragmentAbsorber);
            List<object> fragments = [];

            foreach (var fragment in textFragmentAbsorber.TextFragments)
                fragments.Add(new
                {
                    text = fragment.Text,
                    fontName = fragment.TextState.Font.FontName,
                    fontSize = fragment.TextState.FontSize
                });

            var result = new
            {
                pageIndex,
                totalPages = document.Pages.Count,
                fragmentCount = fragments.Count,
                fragments
            };

            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        }
        else
        {
            var result = new
            {
                pageIndex,
                totalPages = document.Pages.Count,
                text = textAbsorber.Text
            };

            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        }
    }
}