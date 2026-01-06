using System.ComponentModel;
using System.Text.Json;
using Aspose.Words;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Unified tool for managing Word sections (insert, delete, get info)
///     Merges: WordInsertSectionTool, WordDeleteSectionTool, WordGetSectionsTool, WordGetSectionsInfoTool
/// </summary>
[McpServerToolType]
public class WordSectionTool
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
    ///     Initializes a new instance of the WordSectionTool class
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document operations</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation</param>
    public WordSectionTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
    }

    /// <summary>
    ///     Executes a Word section operation (insert, delete, get).
    /// </summary>
    /// <param name="operation">The operation to perform: insert, delete, get.</param>
    /// <param name="path">Word document file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="sectionBreakType">Section break type: NextPage, Continuous, EvenPage, OddPage (for insert).</param>
    /// <param name="insertAtParagraphIndex">Paragraph index to insert section break after (0-based, for insert).</param>
    /// <param name="sectionIndex">Section index (0-based, for insert/delete/get).</param>
    /// <param name="sectionIndices">Array of section indices to delete (0-based, for delete).</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "word_section")]
    [Description(@"Manage Word document sections. Supports 3 operations: insert, delete, get.

Usage examples:
- Insert section: word_section(operation='insert', path='doc.docx', sectionBreakType='NextPage', insertAtParagraphIndex=5)
- Delete section: word_section(operation='delete', path='doc.docx', sectionIndex=1)
- Get sections: word_section(operation='get', path='doc.docx')

Notes:
- Section break types: NextPage (new page), Continuous (same page), EvenPage, OddPage
- IMPORTANT: Deleting a section will also delete all content within that section (paragraphs, tables, images)
- Use 'get' operation first to see section indices and their content statistics before deleting")]
    public string Execute(
        [Description("Operation: insert, delete, get")]
        string operation,
        [Description("Document file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Section break type: NextPage, Continuous, EvenPage, OddPage (for insert)")]
        string? sectionBreakType = null,
        [Description("Paragraph index to insert section break after (0-based, for insert)")]
        int? insertAtParagraphIndex = null,
        [Description("Section index (0-based, for insert/delete/get)")]
        int? sectionIndex = null,
        [Description("Array of section indices to delete (0-based, overrides sectionIndex, for delete)")]
        int[]? sectionIndices = null)
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);

        return operation.ToLower() switch
        {
            "insert" => InsertSection(ctx, outputPath, sectionBreakType, insertAtParagraphIndex, sectionIndex),
            "delete" => DeleteSection(ctx, outputPath, sectionIndex, sectionIndices),
            "get" => GetSections(ctx, sectionIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Inserts a section break into the document at specified position.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sectionBreakType">The type of section break to insert.</param>
    /// <param name="insertAtParagraphIndex">The paragraph index to insert the section break after.</param>
    /// <param name="sectionIndex">The section index containing the target paragraph.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when sectionBreakType is null or empty, or indices are invalid.</exception>
    private static string InsertSection(DocumentContext<Document> ctx, string? outputPath, string? sectionBreakType,
        int? insertAtParagraphIndex, int? sectionIndex)
    {
        if (string.IsNullOrEmpty(sectionBreakType))
            throw new ArgumentException("sectionBreakType is required for insert operation");

        var doc = ctx.Document;
        var builder = new DocumentBuilder(doc);

        var breakType = GetSectionStart(sectionBreakType);

        if (insertAtParagraphIndex.HasValue && insertAtParagraphIndex.Value != -1)
        {
            var actualSectionIndex = sectionIndex ?? 0;
            if (actualSectionIndex < 0 || actualSectionIndex >= doc.Sections.Count)
                throw new ArgumentException(
                    $"sectionIndex must be between 0 and {doc.Sections.Count - 1}, got: {actualSectionIndex}");

            var section = doc.Sections[actualSectionIndex];
            var paragraphs = section.Body.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();

            if (paragraphs.Count == 0)
                throw new ArgumentException("Section has no paragraphs to insert section break after");

            if (insertAtParagraphIndex.Value < 0 || insertAtParagraphIndex.Value >= paragraphs.Count)
                throw new ArgumentException(
                    $"insertAtParagraphIndex must be between 0 and {paragraphs.Count - 1}, got: {insertAtParagraphIndex.Value}");

            builder.MoveTo(paragraphs[insertAtParagraphIndex.Value]);
        }
        else
        {
            builder.MoveToDocumentEnd();
        }

        builder.InsertBreak(BreakType.SectionBreakContinuous);
        builder.CurrentSection.PageSetup.SectionStart = breakType;

        ctx.Save(outputPath);
        var result = $"Section break inserted ({sectionBreakType})\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Converts section break type string to SectionStart enum.
    /// </summary>
    /// <param name="sectionBreakType">The section break type string.</param>
    /// <returns>The corresponding SectionStart enum value.</returns>
    private static SectionStart GetSectionStart(string sectionBreakType)
    {
        return sectionBreakType switch
        {
            "NextPage" => SectionStart.NewPage,
            "Continuous" => SectionStart.Continuous,
            "EvenPage" => SectionStart.EvenPage,
            "OddPage" => SectionStart.OddPage,
            _ => SectionStart.NewPage
        };
    }

    /// <summary>
    ///     Deletes one or more sections from the document (including all content within).
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sectionIndex">The single section index to delete.</param>
    /// <param name="sectionIndices">The array of section indices to delete.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when no section index is provided or document has only one section.</exception>
    private static string DeleteSection(DocumentContext<Document> ctx, string? outputPath, int? sectionIndex,
        int[]? sectionIndices)
    {
        var doc = ctx.Document;
        if (doc.Sections.Count <= 1)
            throw new ArgumentException("Cannot delete the last section. Document must have at least one section.");

        List<int> sectionsToDelete;
        if (sectionIndices is { Length: > 0 })
            sectionsToDelete = sectionIndices.OrderByDescending(s => s).ToList();
        else if (sectionIndex.HasValue)
            sectionsToDelete = [sectionIndex.Value];
        else
            throw new ArgumentException(
                "Either sectionIndex or sectionIndices must be provided for delete operation");

        var deletedCount = 0;
        foreach (var idx in sectionsToDelete)
        {
            if (idx < 0 || idx >= doc.Sections.Count) continue;
            if (doc.Sections.Count <= 1) break;
            doc.Sections.RemoveAt(idx);
            deletedCount++;
        }

        ctx.Save(outputPath);
        var result =
            $"Deleted {deletedCount} section(s) with their content. Remaining sections: {doc.Sections.Count}.\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Gets section information from the document.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="sectionIndex">The specific section index to retrieve, or null for all sections.</param>
    /// <returns>A JSON string containing section information.</returns>
    /// <exception cref="ArgumentException">Thrown when section index is out of range.</exception>
    private static string GetSections(DocumentContext<Document> ctx, int? sectionIndex)
    {
        var doc = ctx.Document;

        if (sectionIndex.HasValue)
        {
            if (sectionIndex.Value < 0 || sectionIndex.Value >= doc.Sections.Count)
                throw new ArgumentException(
                    $"sectionIndex must be between 0 and {doc.Sections.Count - 1}, got: {sectionIndex.Value}");

            var section = doc.Sections[sectionIndex.Value];
            var sectionInfo = BuildSectionInfo(section, sectionIndex.Value);

            var result = new
            {
                totalSections = doc.Sections.Count,
                section = sectionInfo
            };

            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        }
        else
        {
            List<object> sectionList = [];
            for (var i = 0; i < doc.Sections.Count; i++)
            {
                var section = doc.Sections[i];
                sectionList.Add(BuildSectionInfo(section, i));
            }

            var result = new
            {
                totalSections = doc.Sections.Count,
                sections = sectionList
            };

            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        }
    }

    /// <summary>
    ///     Builds section information as a structured object
    /// </summary>
    /// <param name="section">The section to extract information from</param>
    /// <param name="index">The index of the section in the document</param>
    /// <returns>An object containing section information</returns>
    private static object BuildSectionInfo(Section section, int index)
    {
        var pageSetup = section.PageSetup;

        var paragraphs = section.Body.GetChildNodes(NodeType.Paragraph, true);
        var tables = section.Body.GetChildNodes(NodeType.Table, true);
        var shapes = section.Body.GetChildNodes(NodeType.Shape, true);

        var headerCount = 0;
        var footerCount = 0;
        foreach (var hf in section.HeadersFooters.Cast<HeaderFooter>())
            if (hf.HeaderFooterType is HeaderFooterType.HeaderPrimary or HeaderFooterType.HeaderFirst
                or HeaderFooterType.HeaderEven)
            {
                if (!string.IsNullOrWhiteSpace(hf.GetText()))
                    headerCount++;
            }
            else if (hf.HeaderFooterType is HeaderFooterType.FooterPrimary or HeaderFooterType.FooterFirst
                     or HeaderFooterType.FooterEven)
            {
                if (!string.IsNullOrWhiteSpace(hf.GetText()))
                    footerCount++;
            }

        return new
        {
            index,
            sectionBreak = new
            {
                type = GetSectionStartName(pageSetup.SectionStart)
            },
            pageSetup = new
            {
                paperSize = pageSetup.PaperSize.ToString(),
                orientation = pageSetup.Orientation.ToString(),
                margins = new
                {
                    top = pageSetup.TopMargin,
                    bottom = pageSetup.BottomMargin,
                    left = pageSetup.LeftMargin,
                    right = pageSetup.RightMargin
                },
                headerFooterDistance = new
                {
                    header = pageSetup.HeaderDistance,
                    footer = pageSetup.FooterDistance
                },
                pageNumberStart = pageSetup.RestartPageNumbering ? pageSetup.PageStartingNumber : (int?)null,
                differentFirstPage = pageSetup.DifferentFirstPageHeaderFooter,
                differentOddEvenPages = pageSetup.OddAndEvenPagesHeaderFooter,
                columnCount = pageSetup.TextColumns.Count
            },
            contentStatistics = new
            {
                paragraphs = paragraphs.Count,
                tables = tables.Count,
                shapes = shapes.Count
            },
            headersFooters = new
            {
                headers = headerCount,
                footers = footerCount
            }
        };
    }

    /// <summary>
    ///     Converts SectionStart enum to human-readable name
    /// </summary>
    /// <param name="sectionStart">The SectionStart enum value</param>
    /// <returns>Human-readable section break type name</returns>
    private static string GetSectionStartName(SectionStart sectionStart)
    {
        return sectionStart switch
        {
            SectionStart.NewPage => "NextPage",
            SectionStart.Continuous => "Continuous",
            SectionStart.EvenPage => "EvenPage",
            SectionStart.OddPage => "OddPage",
            SectionStart.NewColumn => "NewColumn",
            _ => sectionStart.ToString()
        };
    }
}