using System.ComponentModel;
using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Fields;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Tool for managing cross-references in Word documents
/// </summary>
[McpServerToolType]
public class WordReferenceTool
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
    ///     Initializes a new instance of the WordReferenceTool class
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document operations</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation</param>
    public WordReferenceTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
    }

    /// <summary>
    ///     Executes a Word reference operation (add_table_of_contents, update_table_of_contents, add_index,
    ///     add_cross_reference).
    /// </summary>
    /// <param name="operation">
    ///     The operation to perform: add_table_of_contents, update_table_of_contents, add_index,
    ///     add_cross_reference.
    /// </param>
    /// <param name="path">Word document file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (optional, defaults to overwrite input).</param>
    /// <param name="position">Insert position: start, end (for add_table_of_contents, default: start).</param>
    /// <param name="title">Table of contents title (for add_table_of_contents).</param>
    /// <param name="maxLevel">Maximum heading level to include (for add_table_of_contents, default: 3).</param>
    /// <param name="hyperlinks">Enable clickable hyperlinks (for add_table_of_contents, default: true).</param>
    /// <param name="pageNumbers">Show page numbers (for add_table_of_contents, default: true).</param>
    /// <param name="rightAlignPageNumbers">Right-align page numbers (for add_table_of_contents, default: true).</param>
    /// <param name="tocIndex">TOC field index (0-based, for update_table_of_contents).</param>
    /// <param name="indexEntries">Array of index entries as JSON string (for add_index).</param>
    /// <param name="insertIndexAtEnd">Insert INDEX field at end of document (for add_index, default: true).</param>
    /// <param name="headingStyle">Heading style for index (for add_index, default: 'Heading 1').</param>
    /// <param name="referenceType">Reference type: Heading, Bookmark, Figure, Table, Equation (for add_cross_reference).</param>
    /// <param name="referenceText">Text to insert before reference (for add_cross_reference).</param>
    /// <param name="targetName">Target name (heading text, bookmark name, etc.) (for add_cross_reference).</param>
    /// <param name="insertAsHyperlink">Insert as hyperlink (for add_cross_reference, default: true).</param>
    /// <param name="includeAboveBelow">Include 'above' or 'below' text (for add_cross_reference, default: false).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "word_reference")]
    [Description(
        @"Manage references in Word documents. Supports 4 operations: add_table_of_contents, update_table_of_contents, add_index, add_cross_reference.

Usage examples:
- Add table of contents: word_reference(operation='add_table_of_contents', path='doc.docx', title='Table of Contents', maxLevel=3)
- Update table of contents: word_reference(operation='update_table_of_contents', path='doc.docx')
- Add index: word_reference(operation='add_index', path='doc.docx', indexEntries='[{""text"":""Index term""}]')
- Add cross-reference: word_reference(operation='add_cross_reference', path='doc.docx', referenceType='Bookmark', targetName='Chapter1', referenceText='See ')

Notes:
- TOC is automatically updated after insertion using UpdateFields()
- For cross-references, targetName must be an existing bookmark name in the document
- If headingStyle doesn't exist in the document, it falls back to 'Heading 1'")]
    public string Execute(
        [Description("Operation: add_table_of_contents, update_table_of_contents, add_index, add_cross_reference")]
        string operation,
        [Description("Document file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (optional, defaults to overwrite input)")]
        string? outputPath = null,
        [Description("Insert position: start, end (for add_table_of_contents, default: start)")]
        string position = "start",
        [Description("Table of contents title (for add_table_of_contents, default: 'Table of Contents')")]
        string title = "Table of Contents",
        [Description("Maximum heading level to include (for add_table_of_contents, default: 3)")]
        int maxLevel = 3,
        [Description("Enable clickable hyperlinks (for add_table_of_contents, default: true)")]
        bool hyperlinks = true,
        [Description("Show page numbers (for add_table_of_contents, default: true)")]
        bool pageNumbers = true,
        [Description("Right-align page numbers (for add_table_of_contents, default: true)")]
        bool rightAlignPageNumbers = true,
        [Description("TOC field index (0-based, for update_table_of_contents, optional)")]
        int? tocIndex = null,
        [Description("Array of index entries as JSON string (for add_index)")]
        string? indexEntries = null,
        [Description("Insert INDEX field at end of document (for add_index, default: true)")]
        bool insertIndexAtEnd = true,
        [Description("Heading style for index (for add_index, default: 'Heading 1')")]
        string headingStyle = "Heading 1",
        [Description("Reference type: Heading, Bookmark, Figure, Table, Equation (for add_cross_reference)")]
        string? referenceType = null,
        [Description("Text to insert before reference (for add_cross_reference, optional)")]
        string? referenceText = null,
        [Description("Target name (heading text, bookmark name, etc.) (for add_cross_reference)")]
        string? targetName = null,
        [Description("Insert as hyperlink (for add_cross_reference, default: true)")]
        bool insertAsHyperlink = true,
        [Description("Include 'above' or 'below' text (for add_cross_reference, default: false)")]
        bool includeAboveBelow = false)
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);

        return operation.ToLower() switch
        {
            "add_table_of_contents" => AddTableOfContents(ctx, outputPath, position, title, maxLevel, hyperlinks,
                pageNumbers, rightAlignPageNumbers),
            "update_table_of_contents" => UpdateTableOfContents(ctx, outputPath, tocIndex),
            "add_index" => AddIndex(ctx, outputPath, indexEntries, insertIndexAtEnd, headingStyle),
            "add_cross_reference" => AddCrossReference(ctx, outputPath, referenceType, referenceText, targetName,
                insertAsHyperlink, includeAboveBelow),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds a table of contents to the document.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="position">The insert position: start or end.</param>
    /// <param name="title">The title for the table of contents.</param>
    /// <param name="maxLevel">The maximum heading level to include.</param>
    /// <param name="hyperlinks">Whether to enable clickable hyperlinks.</param>
    /// <param name="pageNumbers">Whether to show page numbers.</param>
    /// <param name="rightAlignPageNumbers">Whether to right-align page numbers.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    private static string AddTableOfContents(
        DocumentContext<Document> ctx,
        string? outputPath,
        string position,
        string title,
        int maxLevel,
        bool hyperlinks,
        bool pageNumbers,
        bool rightAlignPageNumbers)
    {
        var doc = ctx.Document;
        var builder = new DocumentBuilder(doc);

        if (position == "end")
            builder.MoveToDocumentEnd();
        else
            builder.MoveToDocumentStart();

        if (!string.IsNullOrEmpty(title))
        {
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln(title);
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        }

        // Build TOC switches
        var switches = $"\\o \"1-{maxLevel}\"";

        if (!hyperlinks)
            switches += " \\n";

        if (!pageNumbers)
            switches += " \\p \"\"";

        if (!rightAlignPageNumbers)
            switches += " \\l";

        // Use InsertTableOfContents for clearer semantics
        builder.InsertTableOfContents(switches);

        // Update fields to populate TOC content immediately
        doc.UpdateFields();

        ctx.Save(outputPath);

        var result = "Table of contents added\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Updates the table of contents fields in the document.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="tocIndex">The optional zero-based index of a specific TOC field to update.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when tocIndex is out of range.</exception>
    private static string UpdateTableOfContents(
        DocumentContext<Document> ctx,
        string? outputPath,
        int? tocIndex)
    {
        var doc = ctx.Document;
        var tocFields = doc.Range.Fields
            .Where(f => f.Type == FieldType.FieldTOC)
            .ToList();

        if (tocFields.Count == 0)
        {
            var allFields = doc.Range.Fields.ToList();
            var fieldTypes = allFields.Select(f => f.Type.ToString()).Distinct().ToList();
            var message = "No table of contents fields found in document.";
            if (allFields.Count > 0)
                message += $" Found {allFields.Count} field(s) of other types: {string.Join(", ", fieldTypes)}.";
            message += " Use 'add_table_of_contents' operation to add a table of contents first.";
            return message;
        }

        if (tocIndex.HasValue)
        {
            if (tocIndex.Value < 0 || tocIndex.Value >= tocFields.Count)
                throw new ArgumentException($"tocIndex must be between 0 and {tocFields.Count - 1}");
            tocFields[tocIndex.Value].Update();
        }
        else
        {
            foreach (var tocField in tocFields)
                tocField.Update();
        }

        doc.UpdateFields();
        ctx.Save(outputPath);

        var updatedCount = tocIndex.HasValue ? 1 : tocFields.Count;
        var result = $"Updated {updatedCount} table of contents field(s)\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Adds index entries and optionally an INDEX field to the document.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="indexEntriesJson">The index entries as a JSON array string.</param>
    /// <param name="insertIndexAtEnd">Whether to insert an INDEX field at the end of the document.</param>
    /// <param name="headingStyle">The heading style for the index title.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when indexEntriesJson is null, empty, or invalid JSON.</exception>
    private static string AddIndex(
        DocumentContext<Document> ctx,
        string? outputPath,
        string? indexEntriesJson,
        bool insertIndexAtEnd,
        string headingStyle)
    {
        if (string.IsNullOrEmpty(indexEntriesJson))
            throw new ArgumentException("indexEntries is required for add_index operation");

        var indexEntriesArray = JsonNode.Parse(indexEntriesJson)?.AsArray()
                                ?? throw new ArgumentException("indexEntries must be a valid JSON array");

        var doc = ctx.Document;
        var builder = new DocumentBuilder(doc);

        foreach (var entryObj in indexEntriesArray)
            if (entryObj is JsonObject entry)
            {
                var text = entry["text"]?.GetValue<string>();
                var subEntry = entry["subEntry"]?.GetValue<string>();
                var pageRangeBookmark = entry["pageRangeBookmark"]?.GetValue<string>();

                if (!string.IsNullOrEmpty(text))
                {
                    builder.MoveToDocumentEnd();
                    var xeField = $"XE \"{text}\"";
                    if (!string.IsNullOrEmpty(subEntry))
                        xeField += $" \\t \"{subEntry}\"";
                    if (!string.IsNullOrEmpty(pageRangeBookmark))
                        xeField += $" \\r \"{pageRangeBookmark}\"";
                    builder.InsertField(xeField);
                }
            }

        if (insertIndexAtEnd)
        {
            builder.MoveToDocumentEnd();
            builder.InsertBreak(BreakType.PageBreak);

            // Check if headingStyle exists, fallback to Heading1 if not found
            var style = doc.Styles[headingStyle];
            if (style == null)
                builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            else
                builder.ParagraphFormat.Style = style;

            builder.Writeln("Index");
            builder.ParagraphFormat.Style = doc.Styles["Normal"];
            builder.InsertField("INDEX \\e \" \" \\h \"A\"");
        }

        ctx.Save(outputPath);

        var result = $"Index entries added. Total entries: {indexEntriesArray.Count}\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Adds a cross-reference (REF field) to the document.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="referenceType">The reference type: Heading, Bookmark, Figure, Table, or Equation.</param>
    /// <param name="referenceText">The optional text to insert before the reference.</param>
    /// <param name="targetName">The target name (heading text, bookmark name, etc.).</param>
    /// <param name="insertAsHyperlink">Whether to insert the reference as a hyperlink.</param>
    /// <param name="includeAboveBelow">Whether to include 'above' or 'below' text.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when referenceType or targetName is not provided, or referenceType is
    ///     invalid.
    /// </exception>
    private static string AddCrossReference(
        DocumentContext<Document> ctx,
        string? outputPath,
        string? referenceType,
        string? referenceText,
        string? targetName,
        bool insertAsHyperlink,
        bool includeAboveBelow)
    {
        if (string.IsNullOrEmpty(referenceType))
            throw new ArgumentException("referenceType is required for add_cross_reference operation");
        if (string.IsNullOrEmpty(targetName))
            throw new ArgumentException("targetName is required for add_cross_reference operation");

        var validTypes = new[] { "Heading", "Bookmark", "Figure", "Table", "Equation" };
        if (!validTypes.Contains(referenceType, StringComparer.OrdinalIgnoreCase))
            throw new ArgumentException(
                $"Invalid referenceType: {referenceType}. Valid types are: {string.Join(", ", validTypes)}");

        var doc = ctx.Document;
        var builder = new DocumentBuilder(doc);
        builder.MoveToDocumentEnd();

        if (!string.IsNullOrEmpty(referenceText))
            builder.Write(referenceText);

        var fieldCode = insertAsHyperlink ? $"REF {targetName} \\h" : $"REF {targetName}";

        builder.InsertField(fieldCode);
        if (includeAboveBelow)
            builder.Write(" (above)");

        ctx.Save(outputPath);

        var result = $"Cross-reference added (Type: {referenceType})\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }
}