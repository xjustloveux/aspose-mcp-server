using System.Diagnostics.CodeAnalysis;
using System.Drawing.Imaging;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;
using Aspose.Pdf;
using Aspose.Pdf.Devices;
using Aspose.Pdf.Text;
using Aspose.Slides;
using Aspose.Slides.Export;
using Aspose.Words;
using AsposeMcpServer.Core.Progress;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Helpers;
using ModelContextProtocol;
using TxtSaveOptions = Aspose.Cells.TxtSaveOptions;
using Document = Aspose.Words.Document;
using Encoder = System.Drawing.Imaging.Encoder;
using HtmlLoadOptions = Aspose.Pdf.HtmlLoadOptions;
using HtmlSaveOptions = Aspose.Cells.HtmlSaveOptions;
using WordHtmlSaveOptions = Aspose.Words.Saving.HtmlSaveOptions;
using ImageSaveOptions = Aspose.Words.Saving.ImageSaveOptions;
using PageSet = Aspose.Words.Saving.PageSet;
using PdfCompliance = Aspose.Words.Saving.PdfCompliance;
using PdfSaveOptions = Aspose.Words.Saving.PdfSaveOptions;
using SaveFormat = Aspose.Cells.SaveFormat;
using SvgSaveOptions = Aspose.Pdf.SvgSaveOptions;
using WordSaveFormat = Aspose.Words.SaveFormat;

namespace AsposeMcpServer.Core.Conversion;

/// <summary>
///     Options for document conversion operations.
/// </summary>
public class ConversionOptions
{
    /// <summary>
    ///     Optional 1-based page/sheet index for single page/sheet output.
    /// </summary>
    public int? PageIndex { get; init; }

    /// <summary>
    ///     Resolution in DPI for image output. Default is 150.
    /// </summary>
    public int Dpi { get; init; } = 150;

    /// <summary>
    ///     Whether to embed images as Base64 in HTML output. Default is true.
    /// </summary>
    public bool HtmlEmbedImages { get; init; } = true;

    /// <summary>
    ///     Whether to export as single HTML file without external resources. Default is true.
    /// </summary>
    public bool HtmlSingleFile { get; init; } = true;

    /// <summary>
    ///     JPEG quality (1-100) for JPEG image output. Default is 90.
    /// </summary>
    public int JpegQuality { get; init; } = 90;

    /// <summary>
    ///     CSV field separator character. Default is comma.
    /// </summary>
    public string CsvSeparator { get; init; } = ",";

    /// <summary>
    ///     PDF/A compliance level for Word documents: "PDFA1A", "PDFA1B", "PDFA2A", "PDFA2U", "PDFA4".
    ///     Null means no specific compliance.
    /// </summary>
    public string? PdfCompliance { get; init; }

    /// <summary>
    ///     The allowlist of base paths used for symlink resolution before every filesystem sink.
    ///     Populated from <c>ServerConfig.AllowedBasePaths</c> by the caller.
    ///     An empty list disables allowlist enforcement (open-world mode).
    /// </summary>
    public IReadOnlyList<string> AllowedBasePaths { get; init; } = [];

    /// <summary>
    ///     The temp directory of the host performing this conversion, which is where its publish
    ///     records live.
    /// </summary>
    /// <remarks>
    ///     Carried with the options rather than read from a process-wide property, so two hosts in
    ///     one process do not write each other's journals (R18-ARCH01). Defaults to the process
    ///     temp directory for a conversion started outside a host.
    /// </remarks>
    /// <remarks>
    ///     Required, not defaulted. It defaulted to the system temp root while
    ///     <c>CleanupDebtService</c> built its context from the configured session temp
    ///     directory, so on any host with a temp root of its own a conversion journalled where
    ///     nothing would ever recover it — and every call site compiled (R19-REC07).
    /// </remarks>
    public required string RecoveryDirectory { get; init; }

    /// <summary>Where this conversion's publish records go, and the key they are signed with.</summary>
    public RecoveryContext Recovery => RecoveryContext.For(RecoveryDirectory);

    /// <summary>
    ///     Whether an MHT archive may reference resources it does not contain. Converting such a
    ///     file makes the server fetch them, which reaches hosts the caller cannot otherwise
    ///     address. Aspose.Pdf 23.10.0 offers no hook to intercept that for MHT, so the reference
    ///     is refused before conversion unless this is set.
    /// </summary>
    public bool AllowExternalResources { get; init; }

    /// <summary>
    ///     Most model elements the source document may hold for an in-memory conversion, per format.
    ///     Every default is set rather than <c>null</c>; a caller may raise or disable any of them
    ///     by passing its own <see cref="InMemoryModelLimits" />.
    ///     <para>
    ///         Measured (§19.10.3): peak managed memory during an in-memory conversion tracks the
    ///         <em>decompressed</em> document, not the result and not the file on disk. The same
    ///         document shape ran 2,452x, 2,238x and 2,561x its compressed input at three sizes, so
    ///         <see cref="RenderBudget.MaxInMemoryOutputBytes" /> (what is handed back) and
    ///         <c>SessionConfig.MaxFileSizeMb</c> (what is read) both miss the thing that decides
    ///         residency.
    ///     </para>
    ///     <para>
    ///         That measurement says which quantity to bound. It does not say where to put the
    ///         bound, and the numbers that were put there come from a heuristic —
    ///         <see cref="InMemoryModelLimits" /> sets out what was measured for each format and,
    ///         more importantly, what the resulting figure does not cover. Read it before quoting
    ///         a limit as a memory guarantee, because it is not one. This paragraph claimed the
    ///         defaults were "measured, not guessed" while the type it points at called them a
    ///         heuristic that does not bound memory at all; the claim was written when the
    ///         defaults were first set (R9-C02) and not revisited when they were downgraded
    ///         (§23.7), leaving one file saying both things — which is the failure the sentence
    ///         before it was warning about (R13-DOC01).
    ///     </para>
    /// </summary>
    public InMemoryModelLimits ModelLimits { get; init; } = new();

    /// <summary>Options for a conversion with no host behind it.</summary>
    /// <returns>Options whose publish records go to the system temp root.</returns>
    /// <remarks>
    ///     The honest name for what the old default did silently. These public APIs can be called
    ///     without options at all, and nothing at that layer knows where a host keeps its records
    ///     — but saying so at each such site is what keeps the next one from inheriting the wrong
    ///     answer by accident, which is how a conversion came to journal where nothing would ever
    ///     read it (R19-REC07).
    /// </remarks>
    [SuppressMessage("Security", "S5443:Using publicly writable directories is security-sensitive",
        Justification =
            "RecoveryContext.For appends a private .aspose-recovery directory, enforces owner-only protection, and fails closed before any state is written.")]
    public static ConversionOptions WithoutAHost()
    {
        return new ConversionOptions { RecoveryDirectory = Path.GetTempPath() };
    }
}

/// <summary>
///     Ceilings on the source document of an in-memory conversion, one per format because one unit
///     cannot describe four different document models. <c>null</c> disables that format's limit.
///     <para>
///         The defaults were chosen from a <em>heuristic</em>: the increase in managed heap during
///         a conversion of one rich fixture per format, divided by the units in that fixture. That
///         is not the cost of a conversion — it excludes the source model itself, Aspose's native
///         allocations and the process working set — so these limits do not bound how much memory
///         a conversion uses, and nothing here claims they do (§23.7).
///     </para>
///     <para>
///         What the numbers did settle is <em>relative</em>: an earlier round priced the cheapest
///         shape of each format and warned that richer content would cost more. It does, by
///         factors that differ by more than an order of magnitude between formats, so a single
///         safety margin could not stand in for that. Each limit is derived from its own
///         expensive fixture instead (§19.10.3, R9-C02).
///     </para>
///     <list type="bullet">
///         <item>
///             Word — 19,504 nodes of four-cell tables added 52 MB of heap: ~2.7 KB per node,
///             against ~2.5 KB for plain paragraphs. Barely more, so 60,000 nodes stays.
///         </item>
///         <item>
///             Excel — 32,000 cells, one formula in every eighth, added 68 MB: ~2.1 KB per
///             cell, against ~0.4 KB for short strings. Over five times more, so the limit came down
///             from 300,000 to 120,000.
///         </item>
///         <item>
///             PowerPoint — 121 slides carrying twelve text shapes each added 49 MB: ~405 KB
///             per slide, against ~9.5 KB for empty ones. Forty times more, so the limit came down
///             from 1,000 to 600.
///         </item>
///         <item>
///             PDF — 300 pages of two text fragments each added 28 MB: ~93 KB per page. 2,000
///             pages stays.
///         </item>
///     </list>
///     <para>
///         Richer content still exists — a slide of photographs, a scanned page — and these limits
///         do not bound it. Nor do they bound the shapes that were measured: the heap delta on one
///         fixture is a comparison, not a budget. What these ceilings really are is a bound on
///         <em>document size</em>, chosen so that the formats stay in proportion to one another.
///     </para>
///     <para>
///         <strong>The file path is not exempt from the cost, only from the limit.</strong> Measured
///         on the same rich documents, converting to a file peaked at 85% of the in-memory figure
///         for Word, 57% for PDF, 51% for PowerPoint and 35% for Excel: the source model is held
///         either way, and what streaming avoids is holding a second full copy of the result. The
///         refusal points at the file path because it costs less, not because it costs nothing
///         (R9-C02).
///     </para>
/// </summary>
/// <param name="WordNodes">
///     Most nodes an Aspose.Words document may hold. 60,000 is about 162 MB on nodes carrying
///     table cells, and about 150 MB on plain paragraphs.
/// </param>
/// <param name="ExcelCells">
///     Most cells, summed across worksheets, an Aspose.Cells workbook may hold. 120,000 is about
///     252 MB on cells carrying formulas, and about 48 MB on plain short strings.
/// </param>
/// <param name="PowerPointSlides">
///     Most slides an Aspose.Slides presentation may hold. 600 is about 243 MB on slides carrying
///     a dozen text shapes each, and far above any ordinary deck.
/// </param>
/// <param name="PdfPages">
///     Most pages an Aspose.Pdf document may hold. 2,000 leaves room for pages far richer than the
///     text-only ones measured.
/// </param>
public sealed record InMemoryModelLimits(
    int? WordNodes = 60_000,
    int? ExcelCells = 120_000,
    int? PowerPointSlides = 600,
    int? PdfPages = 2_000);

/// <summary>
///     Provides document conversion functionality for various Aspose document types.
///     This utility class is shared between Extension system and Conversion tools.
/// </summary>
public static class DocumentConverter
{
    private const string MhtmlFormat = "mhtml";
    private const string ConvertedDocumentDescription = "converted document";

    /// <summary>
    ///     MIME type mappings for output formats.
    /// </summary>
    private static readonly Dictionary<string, string> MimeTypes = new(StringComparer.OrdinalIgnoreCase)
    {
        { "pdf", "application/pdf" },
        { "html", "text/html" },
        { "htm", "text/html" },
        { "png", "image/png" },
        { "jpg", "image/jpeg" },
        { "jpeg", "image/jpeg" },
        { "tiff", "image/tiff" },
        { "tif", "image/tiff" },
        { "docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document" },
        { "doc", "application/msword" },
        { "rtf", "application/rtf" },
        { "txt", "text/plain" },
        { "odt", "application/vnd.oasis.opendocument.text" },
        { "xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
        { "xls", "application/vnd.ms-excel" },
        { "csv", "text/csv" },
        { "ods", "application/vnd.oasis.opendocument.spreadsheet" },
        { "pptx", "application/vnd.openxmlformats-officedocument.presentationml.presentation" },
        { "ppt", "application/vnd.ms-powerpoint" },
        { "odp", "application/vnd.oasis.opendocument.presentation" },
        { "epub", "application/epub+zip" },
        { "svg", "image/svg+xml" },
        { "xps", "application/vnd.ms-xpsdocument" },
        { "xml", "application/xml" },
        { "md", "text/markdown" },
        { "tex", "application/x-tex" },
        { "mht", "message/rfc822" },
        { MhtmlFormat, "message/rfc822" }
    };

    /// <summary>
    ///     Supported output formats for each document type.
    /// </summary>
    private static readonly Dictionary<DocumentType, HashSet<string>> SupportedFormats =
        new()
        {
            {
                DocumentType.Word,
                new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                {
                    "pdf", "html", "docx", "doc", "rtf", "txt", "odt", "png", "jpg", "jpeg", "tiff", "tif", "bmp", "svg"
                }
            },
            {
                DocumentType.Excel,
                new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                    { "pdf", "html", "xlsx", "xls", "csv", "ods", "png", "jpg", "jpeg", "tiff", "tif", "bmp", "svg" }
            },
            {
                DocumentType.PowerPoint,
                new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                    { "pdf", "html", "pptx", "ppt", "odp", "png", "jpg" }
            },
            {
                DocumentType.Pdf,
                new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                {
                    "docx", "doc", "html", "xlsx", "pptx", "png", "jpg", "jpeg", "tiff", "tif", "epub", "svg", "xps",
                    "xml", "txt"
                }
            }
        };

    /// <summary>
    ///     Resolves symlinks in <paramref name="outputPath" /> and validates the result against
    ///     <paramref name="allowedBasePaths" /> immediately before a filesystem sink.
    ///     When <paramref name="allowedBasePaths" /> is empty (allowlist disabled) the method
    ///     still resolves symlinks and returns the resolved path without an allowlist check —
    ///     consistent with the behaviour of <see cref="SecurityHelper.ResolveAndEnsureWithinAllowlist" />.
    /// </summary>
    /// <param name="outputPath">The user-supplied output file path to resolve.</param>
    /// <param name="allowedBasePaths">
    ///     The allowlist of allowed base paths from <c>ConversionOptions.AllowedBasePaths</c>.
    ///     Pass <see cref="Array.Empty{T}" /> to disable enforcement.
    /// </param>
    /// <returns>The resolved (symlink-free) output path that should be passed to the sink.</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when the resolved path lies outside the allowlist or when a circular
    ///     symlink is encountered.
    /// </exception>
    private static string ResolveOutputPath(string outputPath, IReadOnlyList<string> allowedBasePaths)
    {
        return SecurityHelper.ResolveAndEnsureWithinAllowlist(outputPath, allowedBasePaths, nameof(outputPath));
    }

    #region Document Type Detection

    /// <summary>
    ///     Determines whether the specified extension represents a Word document.
    /// </summary>
    /// <param name="extension">The file extension to check (with or without leading dot).</param>
    /// <returns><c>true</c> if the extension is a Word document format; otherwise, <c>false</c>.</returns>
    public static bool IsWordDocument(string extension)
    {
        var ext = NormalizeExtension(extension);
        return ext is "doc" or "docx" or "rtf" or "odt" or "txt";
    }

    /// <summary>
    ///     Determines whether the specified extension represents an Excel document.
    /// </summary>
    /// <param name="extension">The file extension to check (with or without leading dot).</param>
    /// <returns><c>true</c> if the extension is an Excel document format; otherwise, <c>false</c>.</returns>
    public static bool IsExcelDocument(string extension)
    {
        var ext = NormalizeExtension(extension);
        return ext is "xls" or "xlsx" or "csv" or "ods";
    }

    /// <summary>
    ///     Determines whether the specified extension represents a PowerPoint presentation.
    /// </summary>
    /// <param name="extension">The file extension to check (with or without leading dot).</param>
    /// <returns><c>true</c> if the extension is a PowerPoint format; otherwise, <c>false</c>.</returns>
    public static bool IsPowerPointDocument(string extension)
    {
        var ext = NormalizeExtension(extension);
        return ext is "ppt" or "pptx" or "odp";
    }

    /// <summary>
    ///     Determines whether the specified extension represents a PDF document.
    /// </summary>
    /// <param name="extension">The file extension to check (with or without leading dot).</param>
    /// <returns><c>true</c> if the extension is a PDF document format; otherwise, <c>false</c>.</returns>
    public static bool IsPdfDocument(string extension)
    {
        var ext = NormalizeExtension(extension);
        return ext is "pdf";
    }

    /// <summary>
    ///     Determines whether the specified extension represents an image format.
    /// </summary>
    /// <param name="extension">The file extension to check (with or without leading dot).</param>
    /// <returns><c>true</c> if the extension is an image format; otherwise, <c>false</c>.</returns>
    public static bool IsImageFormat(string extension)
    {
        var ext = NormalizeExtension(extension);
        return ext is "png" or "jpg" or "jpeg" or "tiff" or "tif";
    }

    /// <summary>
    ///     Determines whether the specified format is an image format for Excel conversion.
    /// </summary>
    /// <param name="format">The normalized format string.</param>
    /// <returns><c>true</c> if the format is an image format for Excel; otherwise, <c>false</c>.</returns>
    public static bool IsExcelImageFormat(string format)
    {
        var ext = NormalizeExtension(format);
        return ext is "png" or "jpg" or "jpeg" or "tiff" or "tif" or "bmp" or "svg";
    }

    /// <summary>
    ///     Determines whether the specified extension represents a format that can be converted to PDF.
    ///     Includes HTML, EPUB, Markdown, SVG, XPS, LaTeX, and MHT formats.
    /// </summary>
    /// <param name="extension">The file extension to check (with or without leading dot).</param>
    /// <returns><c>true</c> if the extension can be converted to PDF; otherwise, <c>false</c>.</returns>
    public static bool IsPdfConvertibleFormat(string extension)
    {
        var ext = NormalizeExtension(extension);
        return ext is "html" or "htm" or "epub" or "md" or "svg" or "xps" or "tex" or "mht" or MhtmlFormat;
    }

    /// <summary>
    ///     Gets the document type for the specified file extension.
    /// </summary>
    /// <param name="extension">The file extension to check (with or without leading dot).</param>
    /// <returns>The document type, or <c>null</c> if the extension is not recognized.</returns>
    public static DocumentType? GetDocumentType(string extension)
    {
        if (IsWordDocument(extension)) return DocumentType.Word;
        if (IsExcelDocument(extension)) return DocumentType.Excel;
        if (IsPowerPointDocument(extension)) return DocumentType.PowerPoint;
        if (IsPdfDocument(extension)) return DocumentType.Pdf;
        return null;
    }

    #endregion

    #region Save Format Helpers

    /// <summary>
    ///     Gets the Word save format for the specified format string.
    /// </summary>
    /// <param name="format">The target format (with or without leading dot).</param>
    /// <returns>The corresponding Word save format.</returns>
    /// <exception cref="ArgumentException">Thrown when the format is not supported.</exception>
    public static WordSaveFormat GetWordSaveFormat(string format)
    {
        var ext = NormalizeExtension(format);
        return ext switch
        {
            "pdf" => WordSaveFormat.Pdf,
            "docx" => WordSaveFormat.Docx,
            "doc" => WordSaveFormat.Doc,
            "rtf" => WordSaveFormat.Rtf,
            "html" => WordSaveFormat.Html,
            "txt" => WordSaveFormat.Text,
            "odt" => WordSaveFormat.Odt,
            _ => throw new ArgumentException($"Unsupported output format for Word: {format}")
        };
    }

    /// <summary>
    ///     Gets the Excel save format for the specified format string.
    /// </summary>
    /// <param name="format">The target format (with or without leading dot).</param>
    /// <returns>The corresponding Excel save format.</returns>
    /// <exception cref="ArgumentException">Thrown when the format is not supported.</exception>
    public static SaveFormat GetExcelSaveFormat(string format)
    {
        var ext = NormalizeExtension(format);
        return ext switch
        {
            "pdf" => SaveFormat.Pdf,
            "xlsx" => SaveFormat.Xlsx,
            "xls" => SaveFormat.Excel97To2003,
            "csv" => SaveFormat.Csv,
            "html" => SaveFormat.Html,
            "ods" => SaveFormat.Ods,
            _ => throw new ArgumentException($"Unsupported output format for Excel: {format}")
        };
    }

    /// <summary>
    ///     Gets the PowerPoint save format for the specified format string.
    /// </summary>
    /// <param name="format">The target format (with or without leading dot).</param>
    /// <returns>The corresponding PowerPoint save format.</returns>
    /// <exception cref="ArgumentException">Thrown when the format is not supported.</exception>
    public static Aspose.Slides.Export.SaveFormat GetPresentationSaveFormat(string format)
    {
        var ext = NormalizeExtension(format);
        return ext switch
        {
            "pdf" => Aspose.Slides.Export.SaveFormat.Pdf,
            "pptx" => Aspose.Slides.Export.SaveFormat.Pptx,
            "ppt" => Aspose.Slides.Export.SaveFormat.Ppt,
            "html" => Aspose.Slides.Export.SaveFormat.Html,
            "odp" => Aspose.Slides.Export.SaveFormat.Odp,
            _ => throw new ArgumentException($"Unsupported output format for PowerPoint: {format}")
        };
    }

    /// <summary>
    ///     Gets the PDF save format for the specified format string.
    ///     Plain text ("txt") has no <see cref="Aspose.Pdf.SaveFormat" /> member and is handled
    ///     separately via <see cref="Aspose.Pdf.Text.TextAbsorber" /> in the PDF conversion paths.
    /// </summary>
    /// <param name="format">The target format (with or without leading dot).</param>
    /// <returns>The corresponding PDF save format.</returns>
    /// <exception cref="ArgumentException">Thrown when the format is not supported.</exception>
    public static Aspose.Pdf.SaveFormat GetPdfSaveFormat(string format)
    {
        var ext = NormalizeExtension(format);
        return ext switch
        {
            "docx" => Aspose.Pdf.SaveFormat.DocX,
            "doc" => Aspose.Pdf.SaveFormat.Doc,
            "html" => Aspose.Pdf.SaveFormat.Html,
            "xlsx" => Aspose.Pdf.SaveFormat.Excel,
            "pptx" => Aspose.Pdf.SaveFormat.Pptx,
            "epub" => Aspose.Pdf.SaveFormat.Epub,
            "svg" => Aspose.Pdf.SaveFormat.Svg,
            "xps" => Aspose.Pdf.SaveFormat.Xps,
            "xml" => Aspose.Pdf.SaveFormat.Xml,
            _ => throw new ArgumentException($"Unsupported output format for PDF: {format}")
        };
    }

    #endregion

    #region Stream Conversion (for Extension system)

    /// <summary>
    ///     Reads a presentation's slide count under the gate.
    ///     <para>
    ///         Counting slides enters Aspose.Slides, and this runs before the conversion takes its
    ///         own hold — so on this path the library was touched with nothing held. Found by the
    ///         syntax analyser checking a claim the gate inventory had only asserted by hand
    ///         (§23.30).
    ///     </para>
    /// </summary>
    /// <param name="presentation">The presentation to measure.</param>
    /// <returns>How many slides it holds.</returns>
    private static int SlideCountOf(Presentation presentation)
    {
        using var slidesGate = SlidesGate.Enter();
        return presentation.Slides.Count;
    }

    /// <summary>
    ///     Refuses a source document larger than the operator allowed for an in-memory conversion.
    ///     <para>
    ///         The byte caps bound the result and the file; neither bounds the model, which is what
    ///         peak memory follows (§19.10.3). This is the only place that can look at the model,
    ///         because by the time a conversion starts the document is already loaded — so the
    ///         refusal is about the *next* conversion's memory, not this load's.
    ///     </para>
    ///     <para>
    ///         The ceilings are fixed values on <see cref="InMemoryModelLimits" />, like every
    ///         other <see cref="RenderBudget" /> bound in this server — none of which is a
    ///         configuration setting. A caller constructing its own <see cref="ConversionOptions" />
    ///         can raise or disable them; nothing in this tree does, and there is deliberately no
    ///         operator setting for them (§21.7).
    ///     </para>
    /// </summary>
    /// <param name="document">The loaded source document.</param>
    /// <param name="documentType">Which model it is.</param>
    /// <param name="limits">The ceilings in force; any of them may be null to disable it.</param>
    /// <exception cref="ArgumentException">Thrown when the document is above the configured limit.</exception>
    private static void EnsureModelWithinLimit(object document, DocumentType documentType,
        InMemoryModelLimits limits)
    {
        var (measured, allowed, unit) = documentType switch
        {
            DocumentType.Word when limits.WordNodes.HasValue && document is Document word =>
                (word.GetChildNodes(NodeType.Any, true).Count, limits.WordNodes, "nodes"),
            DocumentType.Excel when limits.ExcelCells.HasValue && document is Workbook workbook =>
                (workbook.Worksheets.Sum(sheet => sheet.Cells.Count), limits.ExcelCells, "cells"),
            DocumentType.PowerPoint when limits.PowerPointSlides.HasValue
                                         && document is Presentation presentation =>
                (SlideCountOf(presentation), limits.PowerPointSlides, "slides"),
            DocumentType.Pdf when limits.PdfPages.HasValue
                                  && document is Aspose.Pdf.Document pdf =>
                (pdf.Pages.Count, limits.PdfPages, "pages"),
            _ => (0, null, string.Empty)
        };

        if (allowed == null || measured <= allowed.Value) return;

        // "Costs less", not "costs nothing": converting the same document to a file still holds
        // the source model, and measured on rich content it peaked at between a third and
        // six-sevenths of the in-memory figure. Saying the file path avoids this would be telling
        // the caller something the measurement does not support (R9-C02).
        throw new ArgumentException(
            $"The document holds {measured:N0} {unit}, above the {allowed.Value:N0} this server "
            + "converts in memory. Convert to a file, which costs less memory, or split the "
            + "document into smaller ones.");
    }

    /// <summary>
    ///     Refuses a file that is already over a limit, before any loader opens it.
    ///     <para>
    ///         <see cref="EnsureModelWithinLimit" /> can only speak once the model exists, so it
    ///         bounds whether to continue rather than what the load costs. Where the file states
    ///         its own size — a presentation package names one part per slide, a PDF's page tree
    ///         carries its count — that number is available for the price of reading a directory
    ///         entry, and the refusal can come first (§23.13.1).
    ///     </para>
    ///     <para>
    ///         Silent when the file cannot answer: a legacy binary, an encrypted package, a PDF
    ///         whose count is not in the tail. Those still reach the loaded-model check. This adds
    ///         an earlier refusal where one is cheap; it does not replace the later one.
    ///     </para>
    /// </summary>
    /// <param name="path">The file about to be opened.</param>
    /// <param name="documentType">Which model it will become.</param>
    /// <param name="limits">The ceilings in force.</param>
    /// <exception cref="ArgumentException">
    ///     Thrown when the file itself says it is above the limit.
    /// </exception>
    public static void EnsureFileWithinLimit(string path, DocumentType? documentType,
        InMemoryModelLimits limits)
    {
        // A presentation that declares more parts than are ever accepted is refused here,
        // whether or not a slide limit is configured: "too large to count" and "could not
        // count" were one null, and null went on to the loader (R22-RES01). One preflight,
        // read once: the refusal and the slide count came from two opens of the path, which
        // need not have been the same bytes (R23-PPT01).
        int? presentationSlides = null;
        if (documentType == DocumentType.PowerPoint)
        {
            var presentation = DocumentSizePreflight.Presentation(path);
            if (presentation.Refusal is { } refusal)
                throw new ArgumentException(refusal, nameof(path));
            presentationSlides = presentation.Slides;
        }

        var (measured, allowed, unit) = documentType switch
        {
            DocumentType.PowerPoint when limits.PowerPointSlides.HasValue =>
                (presentationSlides, limits.PowerPointSlides, "slides"),
            DocumentType.Pdf when limits.PdfPages.HasValue =>
                (DocumentSizePreflight.PageCount(path), limits.PdfPages, "pages"),
            _ => (null, null, string.Empty)
        };

        if (measured == null || allowed == null || measured.Value <= allowed.Value) return;

        throw new ArgumentException(
            $"The file holds {measured.Value:N0} {unit}, above the {allowed.Value:N0} this server "
            + "converts in memory. Convert to a file, which costs less memory, or split the "
            + "document into smaller ones.");
    }

    /// <summary>
    ///     Converts a document to the specified format and returns a Stream.
    /// </summary>
    /// <param name="document">The Aspose document object (Document, Workbook, Presentation, or Aspose.Pdf.Document).</param>
    /// <param name="documentType">The type of the source document.</param>
    /// <param name="outputFormat">The target output format (e.g., "pdf", "html", "png").</param>
    /// <param name="options">Optional conversion options. If null, defaults are used.</param>
    /// <returns>A MemoryStream containing the converted document.</returns>
    /// <exception cref="ArgumentNullException">Thrown when document is null.</exception>
    /// <exception cref="ArgumentException">Thrown when the output format is not supported for the document type.</exception>
    public static Stream ConvertToStream(object document, DocumentType documentType, string outputFormat,
        ConversionOptions? options = null)
    {
        // The model first, while refusing still costs nothing: the byte cap below bounds what is
        // produced, and by then the memory has already been spent (§19.10.3). The limits carry
        // measured defaults; a caller can raise or disable any of them.
        EnsureModelWithinLimit(document, documentType,
            (options ?? ConversionOptions.WithoutAHost()).ModelLimits);

        // The in-memory limit, not the on-disk one: this result is held in memory and usually
        // copied once more on the way out, so the bound that matters is the smaller one (§17.4.2).
        return ConvertToStream(document, documentType, outputFormat,
            RenderBudget.MaxInMemoryOutputBytes, options);
    }

    /// <summary>
    ///     Converts a document to the specified format in memory, under the given byte limit.
    /// </summary>
    /// <param name="document">The Aspose document object.</param>
    /// <param name="documentType">The type of the source document.</param>
    /// <param name="outputFormat">The target output format (e.g., "pdf", "html", "png").</param>
    /// <param name="maxOutputBytes">
    ///     Most bytes the conversion may write. Exposed so a fixture can reach the refusal without
    ///     producing two gigabytes: the production limit is what the public overload passes.
    /// </param>
    /// <param name="options">Optional conversion options. If null, defaults are used.</param>
    /// <returns>A MemoryStream containing the converted document.</returns>
    /// <exception cref="ArgumentNullException">Thrown when document is null.</exception>
    /// <exception cref="ArgumentException">
    ///     Thrown when the output format is not supported for the document type, or the conversion
    ///     writes more than <paramref name="maxOutputBytes" />.
    /// </exception>
    internal static Stream ConvertToStream(object document, DocumentType documentType,
        string outputFormat, long maxOutputBytes, ConversionOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(document);

        if (string.IsNullOrWhiteSpace(outputFormat))
            throw new ArgumentException("Output format cannot be null or empty.", nameof(outputFormat));

        var format = NormalizeExtension(outputFormat);

        if (!IsFormatSupported(documentType, format))
            throw new ArgumentException(
                $"Output format '{format}' is not supported for document type '{documentType}'.");

        var stream = new MemoryStream();

        try
        {
            // Bounded as it is produced. PixelBudget limits the raster dimensions a render is
            // asked for, which says nothing about how many bytes a compressed image or a non-image
            // format then writes, so this path had no limit of its own at all (R4-R01). A refusal
            // throws, and the partial stream goes with it rather than being returned.
            using (var bounded = new BoundedWriteStream(stream, maxOutputBytes, "conversion output"))
            {
                ConvertToStreamInternal(document, documentType, format, bounded, null, options);
            }

            stream.Position = 0;
            return stream;
        }
        catch
        {
            stream.Dispose();
            throw;
        }
    }

    /// <summary>
    ///     Converts a document to the specified format and returns a byte array.
    /// </summary>
    /// <param name="document">The Aspose document object.</param>
    /// <param name="documentType">The type of the source document.</param>
    /// <param name="outputFormat">The target output format.</param>
    /// <param name="options">Optional conversion options. If null, defaults are used.</param>
    /// <returns>A byte array containing the converted document.</returns>
    /// <exception cref="ArgumentNullException">Thrown when document is null.</exception>
    /// <exception cref="ArgumentException">Thrown when the output format is not supported for the document type.</exception>
    public static byte[] ConvertToBytes(object document, DocumentType documentType, string outputFormat,
        ConversionOptions? options = null)
    {
        using var buffer = (MemoryStream)ConvertToStream(document, documentType, outputFormat, options);

        // The buffer is handed over whole when it is exactly the right size, which spares the
        // second full copy that ToArray always makes. Both are bounded by
        // RenderBudget.MaxInMemoryOutputBytes, so the peak is one copy rather than two of an
        // unbounded result (§17.4.2).
        if (buffer.TryGetBuffer(out var segment)
            && segment.Offset == 0 && segment.Count == segment.Array!.Length)
            return segment.Array;

        return buffer.ToArray();
    }

    #endregion

    #region File Conversion (for Tools)

    /// <summary>
    ///     Converts a Word document to the specified output format.
    /// </summary>
    /// <param name="document">The Word document to convert.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="outputFormat">The target output format (with or without leading dot).</param>
    /// <param name="progress">Optional progress reporter (only effective for PDF output).</param>
    /// <param name="options">Optional conversion options. If null, defaults are used.</param>
    /// <returns>
    ///     The paths written, in output order: one file per page for the image formats,
    ///     otherwise the single converted document.
    /// </returns>
    /// <exception cref="ArgumentException">Thrown when the output format is not supported.</exception>
    /// <exception cref="ArgumentOutOfRangeException">Thrown when pageIndex is out of range.</exception>
    public static IReadOnlyList<string> ConvertWordDocument(Document document, string outputPath,
        string outputFormat,
        IProgress<ProgressNotificationValue>? progress = null, ConversionOptions? options = null)
    {
        options ??= ConversionOptions.WithoutAHost();
        document.UpdatePageLayout();
        var format = NormalizeExtension(outputFormat);

        if (IsWordImageFormat(format))
            return ConvertWordToImages(document, outputPath, format, options);

        // H45: resolve symlinks immediately before every write sink (bug 20260415-symlink-toctou-sweep).
        var resolvedOutput = ResolveOutputPath(outputPath, options.AllowedBasePaths);

        if (format == "pdf")
        {
            var saveOptions = new PdfSaveOptions
            {
                ProgressCallback = new WordsProgressAdapter(progress)
            };
            ApplyWordPdfCompliance(saveOptions, options.PdfCompliance);
            BoundedFilePublisher.Publish(resolvedOutput, RenderBudget.MaxOutputBytes,
                stream => document.Save(stream, saveOptions), ConvertedDocumentDescription, options.Recovery,
                options.AllowedBasePaths);
        }
        else if (format is "html" or "htm")
        {
            var saveOptions = new WordHtmlSaveOptions
            {
                ExportImagesAsBase64 = options.HtmlEmbedImages || options.HtmlSingleFile,
                ExportFontsAsBase64 = options.HtmlSingleFile
            };
            BoundedFilePublisher.Publish(resolvedOutput, RenderBudget.MaxOutputBytes,
                stream => document.Save(stream, saveOptions), ConvertedDocumentDescription, options.Recovery,
                options.AllowedBasePaths);
        }
        else
        {
            var saveFormat = GetWordSaveFormat(format);
            BoundedFilePublisher.Publish(resolvedOutput, RenderBudget.MaxOutputBytes,
                stream => document.Save(stream, saveFormat), ConvertedDocumentDescription, options.Recovery,
                options.AllowedBasePaths);
        }

        return [resolvedOutput];
    }

    /// <summary>
    ///     Applies PDF compliance settings to Word PDF save options.
    /// </summary>
    /// <param name="saveOptions">The PDF save options to configure.</param>
    /// <param name="compliance">The compliance string (e.g., "PDFA1A", "PDFA1B").</param>
    private static void ApplyWordPdfCompliance(PdfSaveOptions saveOptions, string? compliance)
    {
        if (string.IsNullOrEmpty(compliance))
            return;

        saveOptions.Compliance = compliance.ToUpperInvariant() switch
        {
            "PDFA1A" => PdfCompliance.PdfA1a,
            "PDFA1B" => PdfCompliance.PdfA1b,
            "PDFA2A" => PdfCompliance.PdfA2a,
            "PDFA2U" => PdfCompliance.PdfA2u,
            "PDFA4" => PdfCompliance.PdfA4,
            _ => saveOptions.Compliance
        };
    }

    /// <summary>
    ///     Determines whether the specified format is an image format for Word conversion.
    /// </summary>
    /// <param name="format">The normalized format string.</param>
    /// <returns><c>true</c> if the format is an image format for Word; otherwise, <c>false</c>.</returns>
    private static bool IsWordImageFormat(string format)
    {
        return format is "png" or "jpg" or "jpeg" or "tiff" or "tif" or "bmp" or "svg";
    }

    /// <summary>
    ///     Converts an Excel workbook to the specified output format.
    /// </summary>
    /// <param name="workbook">The Excel workbook to convert.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="outputFormat">The target output format (with or without leading dot).</param>
    /// <param name="progress">Optional progress reporter (only effective for PDF output).</param>
    /// <param name="options">Optional conversion options. If null, defaults are used.</param>
    /// <returns>
    ///     The paths written, in output order: one file per worksheet for the image formats,
    ///     otherwise the single converted workbook.
    /// </returns>
    /// <exception cref="ArgumentException">Thrown when the output format is not supported.</exception>
    /// <exception cref="ArgumentOutOfRangeException">Thrown when sheetIndex is out of range.</exception>
    public static IReadOnlyList<string> ConvertExcelDocument(Workbook workbook, string outputPath,
        string outputFormat,
        IProgress<ProgressNotificationValue>? progress = null, ConversionOptions? options = null)
    {
        options ??= ConversionOptions.WithoutAHost();
        workbook.CalculateFormula();
        var format = NormalizeExtension(outputFormat);

        if (IsExcelImageFormat(format))
            return ConvertExcelToImages(workbook, outputPath, format, options);

        // H45: resolve symlinks immediately before every write sink (bug 20260415-symlink-toctou-sweep).
        var resolvedOutput = ResolveOutputPath(outputPath, options.AllowedBasePaths);

        if (format == "pdf")
        {
            var saveOptions = new Aspose.Cells.PdfSaveOptions
            {
                PageSavingCallback = new CellsProgressAdapter(progress)
            };
            ApplyExcelPdfCompliance(saveOptions, options.PdfCompliance);
            BoundedFilePublisher.Publish(resolvedOutput, RenderBudget.MaxOutputBytes,
                stream => workbook.Save(stream, saveOptions), ConvertedDocumentDescription, options.Recovery,
                options.AllowedBasePaths);
        }
        else if (format is "html" or "htm")
        {
            var saveOptions = new HtmlSaveOptions
            {
                ExportImagesAsBase64 = options.HtmlEmbedImages || options.HtmlSingleFile,
                SaveAsSingleFile = options.HtmlSingleFile,
                ShowAllSheets = options.HtmlSingleFile
            };

            BoundedFilePublisher.Publish(resolvedOutput, RenderBudget.MaxOutputBytes,
                stream => workbook.Save(stream, saveOptions), ConvertedDocumentDescription, options.Recovery,
                options.AllowedBasePaths);
        }
        else if (format == "csv")
        {
            var saveOptions = new TxtSaveOptions
            {
                Separator = string.IsNullOrEmpty(options.CsvSeparator) ? ',' : options.CsvSeparator[0]
            };
            BoundedFilePublisher.Publish(resolvedOutput, RenderBudget.MaxOutputBytes,
                stream => workbook.Save(stream, saveOptions), ConvertedDocumentDescription, options.Recovery,
                options.AllowedBasePaths);
        }
        else
        {
            var saveFormat = GetExcelSaveFormat(format);
            BoundedFilePublisher.Publish(resolvedOutput, RenderBudget.MaxOutputBytes,
                stream => workbook.Save(stream, saveFormat), ConvertedDocumentDescription, options.Recovery,
                options.AllowedBasePaths);
        }

        return [resolvedOutput];
    }

    /// <summary>
    ///     Applies PDF compliance settings to Excel PDF save options.
    /// </summary>
    /// <param name="saveOptions">The PDF save options to configure.</param>
    /// <param name="compliance">The compliance string (e.g., "PDFA1A", "PDFA1B").</param>
    private static void ApplyExcelPdfCompliance(Aspose.Cells.PdfSaveOptions saveOptions, string? compliance)
    {
        if (string.IsNullOrEmpty(compliance))
            return;

        saveOptions.Compliance = compliance.ToUpperInvariant() switch
        {
            "PDFA1A" => Aspose.Cells.Rendering.PdfCompliance.PdfA1a,
            "PDFA1B" => Aspose.Cells.Rendering.PdfCompliance.PdfA1b,
            _ => saveOptions.Compliance
        };
    }

    /// <summary>
    ///     Converts a PowerPoint presentation to the specified output format.
    /// </summary>
    /// <param name="presentation">The PowerPoint presentation to convert.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="outputFormat">The target output format (with or without leading dot).</param>
    /// <param name="progress">Optional progress reporter (only effective for PDF output).</param>
    /// <param name="options">Optional conversion options. If null, defaults are used.</param>
    /// <returns>
    ///     The paths written, in output order: one file per slide for the image formats,
    ///     otherwise the single converted presentation.
    /// </returns>
    /// <exception cref="ArgumentException">Thrown when the output format is not supported.</exception>
    public static IReadOnlyList<string> ConvertPowerPointDocument(Presentation presentation, string outputPath,
        string outputFormat,
        IProgress<ProgressNotificationValue>? progress = null, ConversionOptions? options = null)
    {
        // Taken here rather than at each caller: this is where the presentation is handed to
        // Aspose.Slides, and a caller that already holds the gate re-enters it (SlidesGate).
        using var slidesGate = SlidesGate.Enter();

        options ??= ConversionOptions.WithoutAHost();
        var format = NormalizeExtension(outputFormat);

        // H45: resolve symlinks immediately before every write sink (bug 20260415-symlink-toctou-sweep).
        var resolvedOutput = ResolveOutputPath(outputPath, options.AllowedBasePaths);

        if (format == "pdf")
        {
            var saveOptions = new PdfOptions
            {
                ProgressCallback = new SlidesProgressAdapter(progress),
                JpegQuality = (byte)options.JpegQuality
            };
            ApplySlidesPdfCompliance(saveOptions, options.PdfCompliance);
            BoundedFilePublisher.Publish(resolvedOutput, RenderBudget.MaxOutputBytes,
                stream => presentation.Save(stream, Aspose.Slides.Export.SaveFormat.Pdf, saveOptions),
                ConvertedDocumentDescription, options.Recovery, options.AllowedBasePaths);
        }
        else
        {
            var saveFormat = GetPresentationSaveFormat(format);
            BoundedFilePublisher.Publish(resolvedOutput, RenderBudget.MaxOutputBytes,
                stream => presentation.Save(stream, saveFormat), ConvertedDocumentDescription, options.Recovery,
                options.AllowedBasePaths);
        }

        return [resolvedOutput];
    }

    /// <summary>
    ///     Applies PDF compliance settings to PowerPoint PDF save options.
    /// </summary>
    /// <param name="saveOptions">The PDF save options to configure.</param>
    /// <param name="compliance">The compliance string (e.g., "PDFA1A", "PDFA1B", "PDFUA").</param>
    private static void ApplySlidesPdfCompliance(PdfOptions saveOptions, string? compliance)
    {
        if (string.IsNullOrEmpty(compliance))
            return;

        saveOptions.Compliance = compliance.ToUpperInvariant() switch
        {
            "PDFA1A" => Aspose.Slides.Export.PdfCompliance.PdfA1a,
            "PDFA1B" => Aspose.Slides.Export.PdfCompliance.PdfA1b,
            "PDFA2A" => Aspose.Slides.Export.PdfCompliance.PdfA2a,
            "PDFA2B" => Aspose.Slides.Export.PdfCompliance.PdfA2b,
            "PDFA3A" => Aspose.Slides.Export.PdfCompliance.PdfA3a,
            "PDFA3B" => Aspose.Slides.Export.PdfCompliance.PdfA3b,
            "PDFUA" => Aspose.Slides.Export.PdfCompliance.PdfUa,
            _ => saveOptions.Compliance
        };
    }

    /// <summary>
    ///     Converts a PDF document to the specified output format.
    ///     For image formats, converts all pages (PNG/JPEG: one file per page, TIFF: single multi-page file).
    /// </summary>
    /// <param name="pdfDocument">The PDF document to convert.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="outputFormat">The target output format (with or without leading dot).</param>
    /// <param name="options">Optional conversion options. If null, defaults are used.</param>
    /// <returns>
    ///     The paths written, in page order: one file per page for PNG, JPEG and BMP, and a
    ///     single file for TIFF and for every non-image format.
    /// </returns>
    /// <exception cref="ArgumentException">Thrown when the output format is not supported.</exception>
    /// <exception cref="ArgumentOutOfRangeException">Thrown when pageIndex is out of range.</exception>
    public static IReadOnlyList<string> ConvertPdfDocument(Aspose.Pdf.Document pdfDocument, string outputPath,
        string outputFormat,
        ConversionOptions? options = null)
    {
        options ??= ConversionOptions.WithoutAHost();
        var format = NormalizeExtension(outputFormat);

        if (IsImageFormat(format))
            return ConvertPdfToImages(pdfDocument, outputPath, format, options.PageIndex, options);

        // H45: resolve symlinks immediately before every write sink (bug 20260415-symlink-toctou-sweep).
        var resolvedOutput = ResolveOutputPath(outputPath, options.AllowedBasePaths);

        if (format is "html" or "htm")
        {
            var htmlOptions = new Aspose.Pdf.HtmlSaveOptions
            {
                PartsEmbeddingMode = Aspose.Pdf.HtmlSaveOptions.PartsEmbeddingModes.EmbedAllIntoHtml,
                RasterImagesSavingMode =
                    Aspose.Pdf.HtmlSaveOptions.RasterImagesSavingModes.AsEmbeddedPartsOfPngPageBackground,
                FontSavingMode = Aspose.Pdf.HtmlSaveOptions.FontSavingModes.SaveInAllFormats
            };
            BoundedFilePublisher.Publish(resolvedOutput, RenderBudget.MaxOutputBytes,
                stream => pdfDocument.Save(stream, htmlOptions), ConvertedDocumentDescription, options.Recovery,
                options.AllowedBasePaths);
        }
        else if (format == "svg")
        {
            var svgOptions = new SvgSaveOptions
            {
                CompressOutputToZipArchive = false
            };
            BoundedFilePublisher.Publish(resolvedOutput, RenderBudget.MaxOutputBytes,
                stream => pdfDocument.Save(stream, svgOptions), ConvertedDocumentDescription, options.Recovery,
                options.AllowedBasePaths);
        }
        else if (format == "txt")
        {
            BoundedFilePublisher.Publish(resolvedOutput, RenderBudget.MaxOutputBytes,
                stream => WritePdfPlainText(stream, pdfDocument),
                ConvertedDocumentDescription, options.Recovery, options.AllowedBasePaths);
        }
        else
        {
            var saveFormat = GetPdfSaveFormat(format);
            BoundedFilePublisher.Publish(resolvedOutput, RenderBudget.MaxOutputBytes,
                stream => pdfDocument.Save(stream, saveFormat), ConvertedDocumentDescription, options.Recovery,
                options.AllowedBasePaths);
        }

        return [resolvedOutput];
    }

    /// <summary>
    ///     Converts a PDF file to images, one file per page (PNG/JPEG) or a single multi-page file (TIFF).
    /// </summary>
    /// <param name="inputPath">The input PDF file path.</param>
    /// <param name="outputPath">The output image file path (page number will be appended for multi-page PNG/JPEG).</param>
    /// <param name="outputFormat">The target image format (with or without leading dot).</param>
    /// <param name="pageIndex">Optional 1-based page index for single page output (omit for all pages).</param>
    /// <param name="options">Optional conversion options. If null, defaults are used.</param>
    /// <returns>
    ///     The image paths written, in page order: one file per page for PNG, JPEG and BMP,
    ///     and a single file for TIFF.
    /// </returns>
    /// <exception cref="ArgumentOutOfRangeException">Thrown when pageIndex is out of range.</exception>
    /// <exception cref="ArgumentException">Thrown when the image format is not supported.</exception>
    public static IReadOnlyList<string> ConvertPdfToImages(string inputPath, string outputPath,
        string outputFormat, int? pageIndex = null, ConversionOptions? options = null)
    {
        using var pdfDoc = new Aspose.Pdf.Document(inputPath);
        return ConvertPdfToImages(pdfDoc, outputPath, outputFormat, pageIndex, options);
    }

    /// <summary>
    ///     Converts a PDF document to images, one file per page (PNG/JPEG) or a single multi-page file (TIFF).
    /// </summary>
    /// <param name="pdfDocument">The PDF document to convert.</param>
    /// <param name="outputPath">The output image file path (page number will be appended for multi-page PNG/JPEG).</param>
    /// <param name="outputFormat">The target image format (with or without leading dot).</param>
    /// <param name="pageIndex">Optional 1-based page index for single page output (omit for all pages).</param>
    /// <param name="options">Optional conversion options. If null, defaults are used.</param>
    /// <returns>
    ///     The image paths written, in page order: one file per page for PNG, JPEG and BMP,
    ///     and a single file for TIFF.
    /// </returns>
    /// <exception cref="ArgumentOutOfRangeException">Thrown when pageIndex is out of range.</exception>
    /// <exception cref="ArgumentException">Thrown when the image format is not supported.</exception>
    public static IReadOnlyList<string> ConvertPdfToImages(Aspose.Pdf.Document pdfDocument, string outputPath,
        string outputFormat, int? pageIndex = null, ConversionOptions? options = null)
    {
        List<string> written = [];
        options ??= ConversionOptions.WithoutAHost();
        var format = NormalizeExtension(outputFormat);

        var ext = Path.GetExtension(outputPath);
        if (string.IsNullOrEmpty(ext))
            throw new ArgumentException(
                $"outputPath must include a file extension (e.g. output.{format}). Received: '{outputPath}'");

        var resolution = new Resolution(options.Dpi);

        if (pageIndex.HasValue)
        {
            if (pageIndex.Value < 1 || pageIndex.Value > pdfDocument.Pages.Count)
                throw new ArgumentOutOfRangeException(nameof(pageIndex),
                    $"Page index must be between 1 and {pdfDocument.Pages.Count}");

            // One page is still a bitmap, and a page can be far larger than A4. This branch had
            // no budget check at all, so a single 200-inch sheet at 3,000 DPI was unbounded
            // (R3-R03).
            AddPdfPageToBudget(new PixelBudget(), pdfDocument.Pages[pageIndex.Value], options.Dpi);

            if (format is "tiff" or "tif")
            {
                var tiffDevice = new TiffDevice(resolution);
                // H46: resolve symlinks immediately before the write sink (bug 20260415-symlink-toctou-sweep).
                var resolvedSingleTiff = ResolveOutputPath(outputPath, options.AllowedBasePaths);
                BoundedFilePublisher.Publish(resolvedSingleTiff, RenderBudget.MaxOutputBytes,
                    stream => tiffDevice.Process(pdfDocument, pageIndex.Value, pageIndex.Value, stream),
                    "rendered page", options.Recovery, options.AllowedBasePaths);
                written.Add(resolvedSingleTiff);
            }
            else
            {
                // H46: resolve symlinks immediately before the write sink (bug 20260415-symlink-toctou-sweep).
                var resolvedSinglePage = ResolveOutputPath(outputPath, options.AllowedBasePaths);
                PageDevice device = format switch
                {
                    "png" => new PngDevice(resolution),
                    "jpg" or "jpeg" => new JpegDevice(resolution, options.JpegQuality),
                    _ => throw new ArgumentException($"Unsupported image format: {format}")
                };

                BoundedFilePublisher.Publish(resolvedSinglePage, RenderBudget.MaxOutputBytes,
                    stream => device.Process(pdfDocument.Pages[pageIndex.Value], stream),
                    "rendered page", options.Recovery, options.AllowedBasePaths);
                written.Add(resolvedSinglePage);
            }

            return written;
        }

        // Every page is rendered from here on, so the cost is the page count times the area
        // at the requested resolution — the product no single parameter limit constrains. The
        // count is checked first because it is cheap; the pixels are then counted from each
        // page's real size rather than from an A4 assumption (R3-R03).
        RenderBudget.EnsureOutputCount(pdfDocument.Pages.Count, "image files");

        var pdfBudget = new PixelBudget();
        foreach (var page in pdfDocument.Pages)
            AddPdfPageToBudget(pdfBudget, page, options.Dpi);

        if (format is "tiff" or "tif")
        {
            var tiffDevice = new TiffDevice(resolution);
            // H46: resolve symlinks immediately before the write sink (bug 20260415-symlink-toctou-sweep).
            var resolvedTiff = ResolveOutputPath(outputPath, options.AllowedBasePaths);
            BoundedFilePublisher.Publish(resolvedTiff, RenderBudget.MaxOutputBytes,
                stream => tiffDevice.Process(pdfDocument, stream),
                "rendered page", options.Recovery, options.AllowedBasePaths);
            written.Add(resolvedTiff);
            return written;
        }

        var dir = Path.GetDirectoryName(outputPath) ?? ".";
        var nameWithoutExt = Path.GetFileNameWithoutExtension(outputPath);

        // One batch for the whole fan-out. Publishing each page on its own handed every page a
        // fresh maximum budget — N pages could produce N times the limit — and put each file at
        // its destination as it was made, so a failure on the last page left the earlier ones
        // behind (R8-C03). Staging them all and publishing once gives the request a single byte
        // budget and an all-or-nothing result.
        using var pageBatch = new BoundedFileBatch(RenderBudget.MaxOutputBytes, "rendered pages",
            options.Recovery, options.AllowedBasePaths);

        for (var i = 1; i <= pdfDocument.Pages.Count; i++)
        {
            var pagePath = pdfDocument.Pages.Count == 1
                ? outputPath
                : Path.Combine(dir, $"{nameWithoutExt}_{i}{ext}");

            // H46: resolve symlinks immediately before the write sink (bug 20260415-symlink-toctou-sweep).
            var resolvedPagePath = ResolveOutputPath(pagePath, options.AllowedBasePaths);
            PageDevice device = format switch
            {
                "png" => new PngDevice(resolution),
                "jpg" or "jpeg" => new JpegDevice(resolution, options.JpegQuality),
                _ => throw new ArgumentException($"Unsupported image format: {format}")
            };

            var page = pdfDocument.Pages[i];
            pageBatch.Stage(resolvedPagePath, stream => device.Process(page, stream));
            written.Add(resolvedPagePath);
        }

        pageBatch.Publish();

        return written;
    }

    /// <summary>
    ///     Converts Word document pages to images.
    /// </summary>
    /// <param name="document">The Word document to convert.</param>
    /// <param name="outputPath">The output file path (page number will be appended for multi-page output).</param>
    /// <param name="outputFormat">The target image format (png, jpg, jpeg, tiff, tif, bmp, svg).</param>
    /// <param name="options">Conversion options including page index and DPI.</param>
    /// <returns>
    ///     The image paths written, in page order — the one requested page, or every page.
    /// </returns>
    /// <exception cref="ArgumentException">Thrown when the format is not supported.</exception>
    /// <exception cref="ArgumentOutOfRangeException">Thrown when pageIndex is out of range.</exception>
    public static IReadOnlyList<string> ConvertWordToImages(Document document, string outputPath,
        string outputFormat,
        ConversionOptions options)
    {
        List<string> written = [];
        var format = NormalizeExtension(outputFormat);
        var saveFormat = format switch
        {
            "png" => WordSaveFormat.Png,
            "jpg" or "jpeg" => WordSaveFormat.Jpeg,
            "tiff" or "tif" => WordSaveFormat.Tiff,
            "bmp" => WordSaveFormat.Bmp,
            "svg" => WordSaveFormat.Svg,
            _ => throw new ArgumentException($"Unsupported image format for Word: {format}")
        };

        if (options.PageIndex.HasValue)
        {
            if (options.PageIndex.Value < 1 || options.PageIndex.Value > document.PageCount)
                throw new ArgumentOutOfRangeException(nameof(options),
                    $"Page index must be between 1 and {document.PageCount}");

            // A single page was never measured, so a very large page at a high resolution was
            // an unbounded allocation (R3-R03).
            AddWordPageToBudget(new PixelBudget(), document, options.PageIndex.Value - 1, options.Dpi);

            var imageOptions = CreateWordImageSaveOptions(saveFormat, options);
            imageOptions.PageSet = new PageSet(options.PageIndex.Value - 1);
            // H45: resolve symlinks immediately before the write sink (bug 20260415-symlink-toctou-sweep).
            var resolvedSinglePage = ResolveOutputPath(outputPath, options.AllowedBasePaths);
            BoundedFilePublisher.Publish(resolvedSinglePage, RenderBudget.MaxOutputBytes,
                stream => document.Save(stream, imageOptions), ConvertedDocumentDescription, options.Recovery,
                options.AllowedBasePaths);
            written.Add(resolvedSinglePage);
        }
        else
        {
            // Every page is rendered here, so the cost is the page count times the area at the
            // requested resolution — the product no single parameter limit constrains.
            RenderBudget.EnsureOutputCount(document.PageCount, "image files");

            var wordBudget = new PixelBudget();
            for (var page = 0; page < document.PageCount; page++)
                AddWordPageToBudget(wordBudget, document, page, options.Dpi);

            var dir = Path.GetDirectoryName(outputPath) ?? ".";
            var baseName = Path.GetFileNameWithoutExtension(outputPath);
            var ext = Path.GetExtension(outputPath);

            using var pageBatch = new BoundedFileBatch(RenderBudget.MaxOutputBytes,
                "rendered pages", options.Recovery, options.AllowedBasePaths);

            for (var i = 0; i < document.PageCount; i++)
            {
                var imageOptions = CreateWordImageSaveOptions(saveFormat, options);
                imageOptions.PageSet = new PageSet(i);
                var pagePath = document.PageCount == 1
                    ? outputPath
                    : Path.Combine(dir, $"{baseName}_{i + 1}{ext}");
                // H45: resolve symlinks immediately before the write sink (bug 20260415-symlink-toctou-sweep).
                var resolvedPagePath = ResolveOutputPath(pagePath, options.AllowedBasePaths);
                pageBatch.Stage(resolvedPagePath, stream => document.Save(stream, imageOptions));
                written.Add(resolvedPagePath);
            }

            pageBatch.Publish();
        }

        return written;
    }

    /// <summary>
    ///     Creates Word image save options with JPEG quality support.
    /// </summary>
    /// <param name="saveFormat">The image save format.</param>
    /// <param name="options">Conversion options.</param>
    /// <returns>Configured image save options.</returns>
    private static ImageSaveOptions CreateWordImageSaveOptions(WordSaveFormat saveFormat, ConversionOptions options)
    {
        var imageOptions = new ImageSaveOptions(saveFormat)
        {
            Resolution = options.Dpi
        };

        if (saveFormat == WordSaveFormat.Jpeg)
            imageOptions.JpegQuality = options.JpegQuality;

        return imageOptions;
    }

    /// <summary>
    ///     Converts Excel workbook sheets to images.
    /// </summary>
    /// <param name="workbook">The Excel workbook to convert.</param>
    /// <param name="outputPath">The output file path (sheet number will be appended for multi-sheet output).</param>
    /// <param name="outputFormat">The target image format (png, jpg, jpeg, tiff, tif, bmp, svg).</param>
    /// <param name="options">Conversion options including sheet index and DPI.</param>
    /// <returns>
    ///     The image paths written, in sheet order — the one requested sheet, or every sheet.
    /// </returns>
    /// <exception cref="ArgumentException">Thrown when the format is not supported.</exception>
    /// <exception cref="ArgumentOutOfRangeException">Thrown when sheetIndex is out of range.</exception>
    public static IReadOnlyList<string> ConvertExcelToImages(Workbook workbook, string outputPath,
        string outputFormat,
        ConversionOptions options)
    {
        List<string> written = [];
        var format = NormalizeExtension(outputFormat);
        var imageType = format switch
        {
            "png" => ImageType.Png,
            "jpg" or "jpeg" => ImageType.Jpeg,
            "tiff" or "tif" => ImageType.Tiff,
            "bmp" => ImageType.Bmp,
            "svg" => ImageType.Svg,
            _ => throw new ArgumentException($"Unsupported image format for Excel: {format}")
        };

        var imageOptions = new ImageOrPrintOptions
        {
            ImageType = imageType,
            HorizontalResolution = options.Dpi,
            VerticalResolution = options.Dpi,
            OnePagePerSheet = true
        };

        if (imageType == ImageType.Jpeg)
            imageOptions.Quality = options.JpegQuality;

        if (options.PageIndex.HasValue)
        {
            if (options.PageIndex.Value < 1 || options.PageIndex.Value > workbook.Worksheets.Count)
                throw new ArgumentOutOfRangeException(nameof(options),
                    $"Sheet index must be between 1 and {workbook.Worksheets.Count}");

            var sheet = workbook.Worksheets[options.PageIndex.Value - 1];
            var sr = new SheetRender(sheet, imageOptions);

            // A worksheet rendered as one page has no size limit of its own, and this branch
            // never measured it (R3-R03).
            AddSheetPageToBudget(new PixelBudget(), sr, 0, options.Dpi);
            // H45: resolve symlinks immediately before the write sink (bug 20260415-symlink-toctou-sweep).
            var resolvedSingleSheet = ResolveOutputPath(outputPath, options.AllowedBasePaths);
            BoundedFilePublisher.Publish(resolvedSingleSheet, RenderBudget.MaxOutputBytes,
                stream => sr.ToImage(0, stream), "rendered sheet", options.Recovery, options.AllowedBasePaths);
            written.Add(resolvedSingleSheet);
        }
        else
        {
            // Every sheet is rendered here; the same product applies. Each sheet's real page
            // size is measured as it is reached, before its bitmap exists (R3-R03).
            RenderBudget.EnsureOutputCount(workbook.Worksheets.Count, "image files");
            var sheetBudget = new PixelBudget();

            var dir = Path.GetDirectoryName(outputPath) ?? ".";
            var baseName = Path.GetFileNameWithoutExtension(outputPath);
            var ext = Path.GetExtension(outputPath);

            using var sheetBatch = new BoundedFileBatch(RenderBudget.MaxOutputBytes,
                "rendered sheets", options.Recovery, options.AllowedBasePaths);

            for (var i = 0; i < workbook.Worksheets.Count; i++)
            {
                var sheet = workbook.Worksheets[i];
                var sr = new SheetRender(sheet, imageOptions);
                AddSheetPageToBudget(sheetBudget, sr, 0, options.Dpi);
                var sheetPath = workbook.Worksheets.Count == 1
                    ? outputPath
                    : Path.Combine(dir, $"{baseName}_{i + 1}{ext}");
                // H45: resolve symlinks immediately before the write sink (bug 20260415-symlink-toctou-sweep).
                var resolvedSheetPath = ResolveOutputPath(sheetPath, options.AllowedBasePaths);
                sheetBatch.Stage(resolvedSheetPath, stream => sr.ToImage(0, stream));
                written.Add(resolvedSheetPath);
            }

            sheetBatch.Publish();
        }

        return written;
    }

    /// <summary>
    ///     Adds one PDF page's real size to a render budget.
    /// </summary>
    /// <param name="budget">The budget to charge.</param>
    /// <param name="page">The page about to be rendered.</param>
    /// <param name="dpi">Resolution the page will be rendered at.</param>
    /// <exception cref="ArgumentException">Thrown when the page takes the render past the budget.</exception>
    private static void AddPdfPageToBudget(PixelBudget budget, Page page, int dpi)
    {
        // Aspose reports page geometry in points, which are 1/72 inch.
        budget.Add(page.Rect.Width / 72.0, page.Rect.Height / 72.0, dpi);
    }

    /// <summary>
    ///     Adds one Word page's real size to a render budget.
    /// </summary>
    /// <param name="budget">The budget to charge.</param>
    /// <param name="document">The document being rendered.</param>
    /// <param name="pageIndex">Zero-based index of the page about to be rendered.</param>
    /// <param name="dpi">Resolution the page will be rendered at.</param>
    /// <exception cref="ArgumentException">Thrown when the page takes the render past the budget.</exception>
    private static void AddWordPageToBudget(PixelBudget budget, Document document, int pageIndex, int dpi)
    {
        var info = document.GetPageInfo(pageIndex);
        budget.Add(info.WidthInPoints / 72.0, info.HeightInPoints / 72.0, dpi);
    }

    /// <summary>
    ///     Adds one rendered worksheet page's real size to a render budget.
    /// </summary>
    /// <param name="budget">The budget to charge.</param>
    /// <param name="render">The sheet renderer holding the page.</param>
    /// <param name="pageIndex">Zero-based index of the page about to be rendered.</param>
    /// <param name="dpi">Resolution the page will be rendered at.</param>
    /// <exception cref="ArgumentException">Thrown when the page takes the render past the budget.</exception>
    private static void AddSheetPageToBudget(PixelBudget budget, SheetRender render, int pageIndex, int dpi)
    {
        // A worksheet with nothing printable renders no page, and the caller reports that itself.
        if (render.PageCount <= pageIndex) return;

        var size = render.GetPageSizeInch(pageIndex);
        budget.Add(size?[0] ?? 0, size?[1] ?? 0, dpi);
    }

    /// <summary>
    ///     Converts a special format file (HTML, EPUB, Markdown, SVG, XPS, LaTeX, MHT) to PDF.
    /// </summary>
    /// <param name="inputPath">The input file path.</param>
    /// <param name="outputPath">The output PDF file path.</param>
    /// <param name="allowedBasePaths">
    ///     The allowlist of base paths used for symlink resolution before the write sink.
    ///     Pass an empty list to skip allowlist enforcement (allowlist disabled).
    /// </param>
    /// <param name="allowExternalResources">
    ///     Accepts that converting an MHT archive may issue outbound requests for resources it
    ///     does not contain. Off by default; the archive is refused instead.
    /// </param>
    /// <returns>The source format name for result reporting.</returns>
    /// <param name="recoveryDirectory">
    ///     The temp directory of the host performing this conversion, which is where its publish
    ///     record goes. Null uses the process temp directory, which is where a record with nobody
    ///     to recover it would sit anyway (R18-ARCH01).
    /// </param>
    /// <exception cref="ArgumentException">
    ///     Thrown when the input format is not supported or when outputPath resolves outside the allowlist.
    /// </exception>
    public static string ConvertToPdfFromSpecialFormat(string inputPath, string outputPath,
        IReadOnlyList<string>? allowedBasePaths = null, bool allowExternalResources = false,
        string? recoveryDirectory = null)
    {
        // Where this conversion's publish record goes. A caller inside a host passes its temp
        // directory; one outside a host gets the process temp directory, which is where a journal
        // with nobody to recover it would sit anyway (R18-ARCH01).
        var recovery = RecoveryContext.For(recoveryDirectory ?? Path.GetTempPath());

        var extension = NormalizeExtension(Path.GetExtension(inputPath));

        // None of these inputs' external fetching can be intercepted by this library version.
        // Measured against the pinned Aspose.Pdf with a localhost probe server: converting HTML,
        // Markdown, SVG and EPUB each fetched the URL the document named, and for HTML an
        // instrumented CustomLoaderOfExternalResources recorded zero invocations while the fetch
        // still happened - the callback below is not consulted for an img src on this path. The
        // only effective control is refusing the input before it is opened, which is what MHT
        // already did and what every other convertible format now does too.
        // One immutable copy, scanned and then parsed. The scanner opened the caller's path and
        // every loader below opened it again, so a file replaced between those two opens was
        // converted without having been scanned — and the comment above says that scan is the only
        // control there is for these formats (R19-CNV01).
        // Admitted on size before it is staged, at the limit the scanner will hold it to: an
        // input over the limit used to be copied whole and refused afterwards (R23-RES01).
        using var authorised = ImmutableInputCopy.Of(inputPath, recovery, allowedBasePaths ?? [],
            extension is "mht" or MhtmlFormat
                ? MhtExternalReferenceScanner.MaxEncodedArchiveBytes
                : MhtExternalReferenceScanner.MaxArchiveBytes);
        var scannedInput = authorised.Path;

        if (extension is "mht" or MhtmlFormat)
            MhtExternalReferenceScanner.EnsureSelfContained(scannedInput, allowExternalResources,
                allowedBasePaths);
        else
            MhtExternalReferenceScanner.EnsureNoRemoteReferences(scannedInput,
                allowExternalResources, allowedBasePaths);
        // H45: resolve symlinks immediately before every write sink (bug 20260415-symlink-toctou-sweep).
        var resolvedOutput = ResolveOutputPath(outputPath, allowedBasePaths ?? []);

        switch (extension)
        {
            case "html":
            case "htm":
                using (var pdfDoc = new Aspose.Pdf.Document(scannedInput,
                           new HtmlLoadOptions
                           {
                               CustomLoaderOfExternalResources =
                                   ExternalResourceGuard.CreatePdfStrategy(allowedBasePaths ?? [])
                           }))
                {
                    BoundedFilePublisher.Publish(resolvedOutput, RenderBudget.MaxOutputBytes,
                        stream => pdfDoc.Save(stream), ConvertedDocumentDescription, recovery, allowedBasePaths ?? []);
                }

                return "HTML";

            case "epub":
                using (var pdfDoc = new Aspose.Pdf.Document(scannedInput, new EpubLoadOptions()))
                {
                    BoundedFilePublisher.Publish(resolvedOutput, RenderBudget.MaxOutputBytes,
                        stream => pdfDoc.Save(stream), ConvertedDocumentDescription, recovery, allowedBasePaths ?? []);
                }

                return "EPUB";

            case "md":
                using (var pdfDoc = new Aspose.Pdf.Document(scannedInput, new MdLoadOptions()))
                {
                    BoundedFilePublisher.Publish(resolvedOutput, RenderBudget.MaxOutputBytes,
                        stream => pdfDoc.Save(stream), ConvertedDocumentDescription, recovery, allowedBasePaths ?? []);
                }

                return "Markdown";

            case "svg":
                using (var pdfDoc = new Aspose.Pdf.Document(scannedInput, new SvgLoadOptions()))
                {
                    BoundedFilePublisher.Publish(resolvedOutput, RenderBudget.MaxOutputBytes,
                        stream => pdfDoc.Save(stream), ConvertedDocumentDescription, recovery, allowedBasePaths ?? []);
                }

                return "SVG";

            case "xps":
                using (var pdfDoc = new Aspose.Pdf.Document(scannedInput, new XpsLoadOptions()))
                {
                    BoundedFilePublisher.Publish(resolvedOutput, RenderBudget.MaxOutputBytes,
                        stream => pdfDoc.Save(stream), ConvertedDocumentDescription, recovery, allowedBasePaths ?? []);
                }

                return "XPS";

            case "tex":
                using (var pdfDoc = new Aspose.Pdf.Document(scannedInput, new TeXLoadOptions()))
                {
                    BoundedFilePublisher.Publish(resolvedOutput, RenderBudget.MaxOutputBytes,
                        stream => pdfDoc.Save(stream), ConvertedDocumentDescription, recovery, allowedBasePaths ?? []);
                }

                return "LaTeX";

            case "mht":
            case MhtmlFormat:
                // Aspose.Pdf 23.10.0 exposes CustomLoaderOfExternalResources on HtmlLoadOptions
                // only, so nothing here can intercept what the library fetches for an MHT archive.
                // The guard at the top of this method is what closes that gap: an archive holding a
                // remote reference never reaches this line unless the caller opted in.
                using (var pdfDoc = new Aspose.Pdf.Document(scannedInput, new MhtLoadOptions()))
                {
                    BoundedFilePublisher.Publish(resolvedOutput, RenderBudget.MaxOutputBytes,
                        stream => pdfDoc.Save(stream), ConvertedDocumentDescription, recovery, allowedBasePaths ?? []);
                }

                return "MHT";

            default:
                throw new ArgumentException($"Unsupported format for PDF conversion: {extension}");
        }
    }

    #endregion

    #region Format Support

    /// <summary>
    ///     Gets the MIME type for the specified output format.
    /// </summary>
    /// <param name="outputFormat">The output format (e.g., "pdf", "html", "png").</param>
    /// <returns>The MIME type string, or "application/octet-stream" if not found.</returns>
    public static string GetMimeType(string outputFormat)
    {
        var format = NormalizeExtension(outputFormat);
        return MimeTypes.GetValueOrDefault(format, "application/octet-stream");
    }

    /// <summary>
    ///     Checks whether the specified output format is supported for the given document type.
    /// </summary>
    /// <param name="documentType">The document type to check.</param>
    /// <param name="outputFormat">The output format to check.</param>
    /// <returns><c>true</c> if the format is supported; otherwise, <c>false</c>.</returns>
    public static bool IsFormatSupported(DocumentType documentType, string outputFormat)
    {
        var format = NormalizeExtension(outputFormat);
        return SupportedFormats.TryGetValue(documentType, out var formats) && formats.Contains(format);
    }

    /// <summary>
    ///     Gets all supported output formats for the specified document type.
    /// </summary>
    /// <param name="documentType">The document type.</param>
    /// <returns>An enumerable of supported format strings.</returns>
    public static IEnumerable<string> GetSupportedFormats(DocumentType documentType)
    {
        return SupportedFormats.TryGetValue(documentType, out var formats)
            ? formats
            : Enumerable.Empty<string>();
    }

    #endregion

    #region Private Methods

    /// <summary>
    ///     Normalizes the format/extension string by removing leading dots and converting to lowercase.
    /// </summary>
    /// <param name="format">The format string to normalize.</param>
    /// <returns>The normalized format string.</returns>
    private static string NormalizeExtension(string format)
    {
        return format.TrimStart('.').ToLowerInvariant();
    }

    /// <summary>
    ///     Internal method that performs the actual conversion to a stream.
    /// </summary>
    /// <param name="document">The source document.</param>
    /// <param name="documentType">The document type.</param>
    /// <param name="format">The normalized output format.</param>
    /// <param name="outputStream">The stream to write the converted document to.</param>
    /// <param name="progress">Optional progress reporter.</param>
    /// <param name="options">Optional conversion options.</param>
    /// <exception cref="ArgumentException">Thrown when the document type is not supported.</exception>
    private static void ConvertToStreamInternal(object document, DocumentType documentType, string format,
        Stream outputStream, IProgress<ProgressNotificationValue>? progress, ConversionOptions? options)
    {
        options ??= ConversionOptions.WithoutAHost();

        switch (documentType)
        {
            case DocumentType.Word:
                ConvertWordToStream((Document)document, format, outputStream, progress, options);
                break;

            case DocumentType.Excel:
                ConvertExcelToStream((Workbook)document, format, outputStream, progress, options);
                break;

            case DocumentType.PowerPoint:
                ConvertPowerPointToStream((Presentation)document, format, outputStream, progress, options);
                break;

            case DocumentType.Pdf:
                ConvertPdfToStream((Aspose.Pdf.Document)document, format, outputStream, options);
                break;

            default:
                throw new ArgumentException($"Unsupported document type: {documentType}");
        }
    }

    /// <summary>
    ///     Converts a Word document to a stream.
    ///     For image formats, renders only the first page.
    /// </summary>
    /// <param name="document">The Word document to convert.</param>
    /// <param name="format">The normalized output format.</param>
    /// <param name="outputStream">The stream to write the converted document to.</param>
    /// <param name="progress">Optional progress reporter.</param>
    /// <param name="options">Conversion options.</param>
    /// <exception cref="ArgumentException">Thrown when the output format is not supported.</exception>
    private static void ConvertWordToStream(Document document, string format, Stream outputStream,
        IProgress<ProgressNotificationValue>? progress, ConversionOptions options)
    {
        document.UpdatePageLayout();

        if (IsWordImageFormat(format))
        {
            var saveFormat = format switch
            {
                "png" => WordSaveFormat.Png,
                "jpg" or "jpeg" => WordSaveFormat.Jpeg,
                "tiff" or "tif" => WordSaveFormat.Tiff,
                "bmp" => WordSaveFormat.Bmp,
                "svg" => WordSaveFormat.Svg,
                _ => throw new ArgumentException($"Unsupported image format for Word: {format}")
            };
            // The file path prices its rasters; this one did not, so the same document rendered
            // through the extension bridge was unbounded (R4-R01).
            if (document.PageCount > 0)
                AddWordPageToBudget(new PixelBudget(), document, 0, options.Dpi);

            var imageOptions = new ImageSaveOptions(saveFormat) { PageSet = new PageSet(0) };
            if (saveFormat == WordSaveFormat.Jpeg)
                imageOptions.JpegQuality = options.JpegQuality;
            document.Save(outputStream, imageOptions);
            return;
        }

        if (format == "pdf")
        {
            var saveOptions = new PdfSaveOptions
            {
                ProgressCallback = new WordsProgressAdapter(progress)
            };
            ApplyWordPdfCompliance(saveOptions, options.PdfCompliance);
            document.Save(outputStream, saveOptions);
        }
        else if (format is "html" or "htm")
        {
            var saveOptions = new WordHtmlSaveOptions
            {
                ExportImagesAsBase64 = options.HtmlEmbedImages || options.HtmlSingleFile,
                ExportFontsAsBase64 = options.HtmlSingleFile
            };
            document.Save(outputStream, saveOptions);
        }
        else
        {
            var saveFormat = GetWordSaveFormat(format);
            document.Save(outputStream, saveFormat);
        }
    }

    /// <summary>
    ///     Converts an Excel workbook to a stream.
    ///     For image formats, renders only the first sheet.
    /// </summary>
    /// <param name="workbook">The Excel workbook to convert.</param>
    /// <param name="format">The normalized output format.</param>
    /// <param name="outputStream">The stream to write the converted document to.</param>
    /// <param name="progress">Optional progress reporter.</param>
    /// <param name="options">Conversion options.</param>
    /// <exception cref="ArgumentException">Thrown when the output format is not supported.</exception>
    private static void ConvertExcelToStream(Workbook workbook, string format, Stream outputStream,
        IProgress<ProgressNotificationValue>? progress, ConversionOptions options)
    {
        workbook.CalculateFormula();

        if (IsExcelImageFormat(format))
        {
            var imageType = format switch
            {
                "png" => ImageType.Png,
                "jpg" or "jpeg" => ImageType.Jpeg,
                "tiff" or "tif" => ImageType.Tiff,
                "bmp" => ImageType.Bmp,
                "svg" => ImageType.Svg,
                _ => throw new ArgumentException($"Unsupported image format for Excel: {format}")
            };

            var imageOptions = new ImageOrPrintOptions
            {
                ImageType = imageType,
                OnePagePerSheet = true
            };
            if (imageType == ImageType.Jpeg)
                imageOptions.Quality = options.JpegQuality;

            var sr = new SheetRender(workbook.Worksheets[0], imageOptions);
            AddSheetPageToBudget(new PixelBudget(), sr, 0, options.Dpi);
            sr.ToImage(0, outputStream);
            return;
        }

        if (format == "pdf")
        {
            var saveOptions = new Aspose.Cells.PdfSaveOptions
            {
                PageSavingCallback = new CellsProgressAdapter(progress)
            };
            ApplyExcelPdfCompliance(saveOptions, options.PdfCompliance);
            workbook.Save(outputStream, saveOptions);
        }
        else if (format is "html" or "htm")
        {
            var saveOptions = new HtmlSaveOptions
            {
                ExportImagesAsBase64 = options.HtmlEmbedImages || options.HtmlSingleFile,
                SaveAsSingleFile = options.HtmlSingleFile,
                ShowAllSheets = options.HtmlSingleFile
            };
            workbook.Save(outputStream, saveOptions);
        }
        else if (format == "csv")
        {
            var saveOptions = new TxtSaveOptions
            {
                Separator = string.IsNullOrEmpty(options.CsvSeparator) ? ',' : options.CsvSeparator[0]
            };
            workbook.Save(outputStream, saveOptions);
        }
        else
        {
            var saveFormat = GetExcelSaveFormat(format);
            workbook.Save(outputStream, saveFormat);
        }
    }

    /// <summary>
    ///     Converts a PowerPoint presentation to a stream.
    /// </summary>
    /// <param name="presentation">The PowerPoint presentation to convert.</param>
    /// <param name="format">The normalized output format.</param>
    /// <param name="outputStream">The stream to write the converted document to.</param>
    /// <param name="progress">Optional progress reporter.</param>
    /// <param name="options">Conversion options.</param>
    /// <exception cref="ArgumentException">Thrown when the output format is not supported.</exception>
    /// <exception cref="InvalidOperationException">Thrown when the presentation has no slides (for image output).</exception>
    private static void ConvertPowerPointToStream(Presentation presentation, string format, Stream outputStream,
        IProgress<ProgressNotificationValue>? progress, ConversionOptions options)
    {
        // The in-memory counterpart of ConvertPowerPointDocument, and the same boundary
        // (SlidesGate).
        using var slidesGate = SlidesGate.Enter();

        if (format is "png" or "jpg" or "jpeg")
        {
            ConvertPresentationToImageStream(presentation, format, outputStream, options);
            return;
        }

        if (format == "pdf")
        {
            var saveOptions = new PdfOptions
            {
                ProgressCallback = new SlidesProgressAdapter(progress),
                JpegQuality = (byte)options.JpegQuality
            };
            ApplySlidesPdfCompliance(saveOptions, options.PdfCompliance);
            presentation.Save(outputStream, Aspose.Slides.Export.SaveFormat.Pdf, saveOptions);
        }
        else
        {
            var saveFormat = GetPresentationSaveFormat(format);
            presentation.Save(outputStream, saveFormat);
        }
    }

    /// <summary>
    ///     Converts a PowerPoint presentation to an image stream.
    ///     For multi-slide presentations, renders the first slide.
    /// </summary>
    /// <param name="presentation">The PowerPoint presentation to convert.</param>
    /// <param name="format">The normalized image format (png, jpg, jpeg).</param>
    /// <param name="outputStream">The stream to write the image to.</param>
    /// <param name="options">Conversion options.</param>
    /// <exception cref="InvalidOperationException">Thrown when the presentation has no slides.</exception>
    /// <exception cref="ArgumentException">Thrown when the image format is not supported.</exception>
    private static void ConvertPresentationToImageStream(Presentation presentation, string format, Stream outputStream,
        ConversionOptions options)
    {
        if (presentation.Slides.Count == 0)
            throw new InvalidOperationException("Presentation has no slides to convert.");

        // Slide size is in points, and a deck can set it to anything (R4-R01).
        new PixelBudget().Add(presentation.SlideSize.Size.Width / 72.0,
            presentation.SlideSize.Size.Height / 72.0, 96);

        var slide = presentation.Slides[0];

        if (format is "jpg" or "jpeg")
        {
            var encoder = ImageCodecInfo.GetImageEncoders().First(c => c.FormatID == ImageFormat.Jpeg.Guid);
            var encoderParams = new EncoderParameters(1);
            encoderParams.Param[0] = new EncoderParameter(Encoder.Quality, options.JpegQuality);

            using var bitmap = slide.GetThumbnail(1f, 1f);
            bitmap.Save(outputStream, encoder, encoderParams);
        }
        else
        {
            var imageFormat = format switch
            {
                "png" => ImageFormat.Png,
                _ => throw new ArgumentException($"Unsupported image format: {format}")
            };

            using var bitmap = slide.GetThumbnail(1f, 1f);
            bitmap.Save(outputStream, imageFormat);
        }
    }

    /// <summary>
    ///     Converts a PDF document to a stream.
    ///     For image formats, renders only the first page.
    /// </summary>
    /// <param name="pdfDocument">The PDF document to convert.</param>
    /// <param name="format">The normalized output format.</param>
    /// <param name="outputStream">The stream to write the converted document to.</param>
    /// <param name="options">Conversion options.</param>
    /// <exception cref="ArgumentException">Thrown when the output format is not supported.</exception>
    /// <exception cref="InvalidOperationException">Thrown when the PDF has no pages (for image output).</exception>
    private static void ConvertPdfToStream(Aspose.Pdf.Document pdfDocument, string format, Stream outputStream,
        ConversionOptions options)
    {
        if (IsImageFormat(format))
        {
            ConvertPdfFirstPageToImageStream(pdfDocument, format, outputStream, options);
        }
        else if (format is "html" or "htm")
        {
            var htmlOptions = new Aspose.Pdf.HtmlSaveOptions
            {
                PartsEmbeddingMode = Aspose.Pdf.HtmlSaveOptions.PartsEmbeddingModes.EmbedAllIntoHtml,
                RasterImagesSavingMode =
                    Aspose.Pdf.HtmlSaveOptions.RasterImagesSavingModes.AsEmbeddedPartsOfPngPageBackground,
                FontSavingMode = Aspose.Pdf.HtmlSaveOptions.FontSavingModes.SaveInAllFormats
            };
            pdfDocument.Save(outputStream, htmlOptions);
        }
        else if (format == "svg")
        {
            var svgOptions = new SvgSaveOptions
            {
                CompressOutputToZipArchive = false
            };
            pdfDocument.Save(outputStream, svgOptions);
        }
        else if (format == "txt")
        {
            WritePdfPlainText(outputStream, pdfDocument);
        }
        else
        {
            var saveFormat = GetPdfSaveFormat(format);
            pdfDocument.Save(outputStream, saveFormat);
        }
    }

    /// <summary>
    ///     Writes the plain text of every page of a PDF document to a stream.
    ///     Plain text has no <see cref="Aspose.Pdf.SaveFormat" /> member, so "txt" output is produced
    ///     via <see cref="Aspose.Pdf.Text.TextAbsorber" /> instead of <c>Document.Save</c>.
    ///     Pending paragraphs are laid out first so unsaved edits are included, matching Save-based paths.
    /// </summary>
    /// <param name="stream">The destination to write the text to.</param>
    /// <param name="pdfDocument">The PDF document to extract text from.</param>
    private static void WritePdfPlainText(Stream stream, Aspose.Pdf.Document pdfDocument)
    {
        pdfDocument.ProcessParagraphs();

        // One page at a time. Absorbing the whole document produced a single string holding every
        // page's text, which was then encoded into a second full-size array before the bounded
        // stream saw a byte of it — two copies of an unbounded document ahead of the limit meant
        // to bound it (R8-C03). The peak is now one page's text, and the cap applies as each page
        // is written rather than after everything has been built.
        var encoder = new UTF8Encoding(false).GetEncoder();
        var buffer = new byte[8192];
        var chars = new char[2048];

        for (var i = 1; i <= pdfDocument.Pages.Count; i++)
        {
            var absorber = new TextAbsorber();
            pdfDocument.Pages[i].Accept(absorber);
            var text = absorber.Text;

            for (var offset = 0; offset < text.Length;)
            {
                var take = Math.Min(chars.Length, text.Length - offset);
                text.CopyTo(offset, chars, 0, take);
                offset += take;

                var flush = offset >= text.Length && i == pdfDocument.Pages.Count;
                var written = encoder.GetBytes(chars, 0, take, buffer, 0, flush);
                stream.Write(buffer, 0, written);
            }
        }
    }

    /// <summary>
    ///     Converts the first page of a PDF document to an image stream.
    /// </summary>
    /// <param name="pdfDocument">The PDF document to convert.</param>
    /// <param name="format">The normalized image format (png, jpg, jpeg, tiff, tif).</param>
    /// <param name="outputStream">The stream to write the image to.</param>
    /// <param name="options">Conversion options.</param>
    /// <exception cref="InvalidOperationException">Thrown when the PDF has no pages.</exception>
    /// <exception cref="ArgumentException">Thrown when the image format is not supported.</exception>
    private static void ConvertPdfFirstPageToImageStream(Aspose.Pdf.Document pdfDocument, string format,
        Stream outputStream, ConversionOptions options)
    {
        if (pdfDocument.Pages.Count == 0)
            throw new InvalidOperationException("PDF document has no pages to convert.");

        var resolution = new Resolution(options.Dpi);

        if (format is "tiff" or "tif")
        {
            // This one renders the whole document, not the first page, so every page is priced.
            var wholeDocument = new PixelBudget();
            foreach (var page in pdfDocument.Pages)
                AddPdfPageToBudget(wholeDocument, page, options.Dpi);

            var tiffDevice = new TiffDevice(resolution);
            tiffDevice.Process(pdfDocument, outputStream);
            return;
        }

        AddPdfPageToBudget(new PixelBudget(), pdfDocument.Pages[1], options.Dpi);

        PageDevice device = format switch
        {
            "png" => new PngDevice(resolution),
            "jpg" or "jpeg" => new JpegDevice(resolution, options.JpegQuality),
            _ => throw new ArgumentException($"Unsupported image format: {format}")
        };

        device.Process(pdfDocument.Pages[1], outputStream);
    }

    #endregion
}
