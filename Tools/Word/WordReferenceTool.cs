using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Fields;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Tool for managing cross-references in Word documents
/// </summary>
public class WordReferenceTool : IAsposeTool
{
    public string Description =>
        @"Manage references in Word documents. Supports 4 operations: add_table_of_contents, update_table_of_contents, add_index, add_cross_reference.

Usage examples:
- Add table of contents: word_reference(operation='add_table_of_contents', path='doc.docx', title='Table of Contents', maxLevel=3)
- Update table of contents: word_reference(operation='update_table_of_contents', path='doc.docx')
- Add index: word_reference(operation='add_index', path='doc.docx', indexEntries=[{'text':'Index term'}])
- Add cross-reference: word_reference(operation='add_cross_reference', path='doc.docx', referenceType='Bookmark', targetName='Chapter1', referenceText='See ')

Notes:
- TOC is automatically updated after insertion using UpdateFields()
- For cross-references, targetName must be an existing bookmark name in the document
- If headingStyle doesn't exist in the document, it falls back to 'Heading 1'";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'add_table_of_contents': Add table of contents (required params: path)
- 'update_table_of_contents': Update table of contents (required params: path)
- 'add_index': Add index (required params: path, entries)
- 'add_cross_reference': Add cross-reference (required params: path, referenceType, targetText, displayText)",
                @enum = new[]
                    { "add_table_of_contents", "update_table_of_contents", "add_index", "add_cross_reference" }
            },
            path = new
            {
                type = "string",
                description = "Document file path (required for all operations)"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, defaults to overwrite input)"
            },
            position = new
            {
                type = "string",
                description = "Insert position: start, end (for add_table_of_contents, default: start)",
                @enum = new[] { "start", "end" }
            },
            title = new
            {
                type = "string",
                description = "Table of contents title (for add_table_of_contents, default: 'Table of Contents')"
            },
            maxLevel = new
            {
                type = "number",
                description = "Maximum heading level to include (for add_table_of_contents, default: 3)"
            },
            hyperlinks = new
            {
                type = "boolean",
                description = "Enable clickable hyperlinks (for add_table_of_contents, default: true)"
            },
            pageNumbers = new
            {
                type = "boolean",
                description = "Show page numbers (for add_table_of_contents, default: true)"
            },
            rightAlignPageNumbers = new
            {
                type = "boolean",
                description = "Right-align page numbers (for add_table_of_contents, default: true)"
            },
            tocIndex = new
            {
                type = "number",
                description = "TOC field index (0-based, for update_table_of_contents, optional)"
            },
            indexEntries = new
            {
                type = "array",
                description = "Array of index entries (for add_index)",
                items = new
                {
                    type = "object",
                    properties = new
                    {
                        text = new { type = "string" },
                        subEntry = new { type = "string" },
                        pageRangeBookmark = new { type = "string" }
                    },
                    required = new[] { "text" }
                }
            },
            insertIndexAtEnd = new
            {
                type = "boolean",
                description = "Insert INDEX field at end of document (for add_index, default: true)"
            },
            headingStyle = new
            {
                type = "string",
                description = "Heading style for index (for add_index, default: 'Heading 1')"
            },
            referenceType = new
            {
                type = "string",
                description = "Reference type: Heading, Bookmark, Figure, Table, Equation (for add_cross_reference)",
                @enum = new[] { "Heading", "Bookmark", "Figure", "Table", "Equation" }
            },
            referenceText = new
            {
                type = "string",
                description = "Text to insert before reference (for add_cross_reference, optional)"
            },
            targetName = new
            {
                type = "string",
                description = "Target name (heading text, bookmark name, etc.) (for add_cross_reference)"
            },
            insertAsHyperlink = new
            {
                type = "boolean",
                description = "Insert as hyperlink (for add_cross_reference, default: true)"
            },
            includeAboveBelow = new
            {
                type = "boolean",
                description = "Include 'above' or 'below' text (for add_cross_reference, default: false)"
            }
        },
        required = new[] { "operation", "path" }
    };

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters</param>
    /// <returns>Result message as a string</returns>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

        return operation.ToLower() switch
        {
            "add_table_of_contents" => await AddTableOfContentsAsync(path, outputPath, arguments),
            "update_table_of_contents" => await UpdateTableOfContentsAsync(path, outputPath, arguments),
            "add_index" => await AddIndexAsync(path, outputPath, arguments),
            "add_cross_reference" => await AddCrossReferenceAsync(path, outputPath, arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds a table of contents to the document
    /// </summary>
    /// <param name="path">Source document file path</param>
    /// <param name="outputPath">Output document file path</param>
    /// <param name="arguments">
    ///     JSON arguments containing position, title, maxLevel, hyperlinks, pageNumbers,
    ///     rightAlignPageNumbers
    /// </param>
    /// <returns>Success message with output path</returns>
    private Task<string> AddTableOfContentsAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var position = ArgumentHelper.GetString(arguments, "position", "start");
            var title = ArgumentHelper.GetString(arguments, "title", "Table of Contents");
            var maxLevel = ArgumentHelper.GetInt(arguments, "maxLevel", 3);
            var hyperlinks = ArgumentHelper.GetBool(arguments, "hyperlinks", true);
            var pageNumbers = ArgumentHelper.GetBool(arguments, "pageNumbers", true);
            var rightAlignPageNumbers = ArgumentHelper.GetBool(arguments, "rightAlignPageNumbers", true);

            var doc = new Document(path);
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

            doc.Save(outputPath);
            return $"Table of contents added: {outputPath}";
        });
    }

    /// <summary>
    ///     Updates the table of contents fields in the document
    /// </summary>
    /// <param name="path">Source document file path</param>
    /// <param name="outputPath">Output document file path</param>
    /// <param name="arguments">JSON arguments containing optional tocIndex</param>
    /// <returns>Success message with number of updated TOC fields, or helpful message if no TOC found</returns>
    private Task<string> UpdateTableOfContentsAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var tocIndex = ArgumentHelper.GetIntNullable(arguments, "tocIndex");

            var doc = new Document(path);
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
            doc.Save(outputPath);
            var updatedCount = tocIndex.HasValue ? 1 : tocFields.Count;
            return $"Updated {updatedCount} table of contents field(s): {outputPath}";
        });
    }

    /// <summary>
    ///     Adds index entries and optionally an INDEX field to the document
    /// </summary>
    /// <param name="path">Source document file path</param>
    /// <param name="outputPath">Output document file path</param>
    /// <param name="arguments">JSON arguments containing indexEntries array, insertIndexAtEnd, headingStyle</param>
    /// <returns>Success message with entry count and output path</returns>
    private Task<string> AddIndexAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var indexEntriesArray = ArgumentHelper.GetArray(arguments, "indexEntries");
            var insertIndexAtEnd = ArgumentHelper.GetBool(arguments, "insertIndexAtEnd", true);
            var headingStyle = ArgumentHelper.GetString(arguments, "headingStyle", "Heading 1");

            var doc = new Document(path);
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

            doc.Save(outputPath);
            return $"Index entries added. Total entries: {indexEntriesArray.Count}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Adds a cross-reference (REF field) to the document
    /// </summary>
    /// <param name="path">Source document file path</param>
    /// <param name="outputPath">Output document file path</param>
    /// <param name="arguments">
    ///     JSON arguments containing referenceType, targetName, referenceText, insertAsHyperlink,
    ///     includeAboveBelow
    /// </param>
    /// <returns>Success message with reference type and output path</returns>
    /// <exception cref="ArgumentException">Thrown when referenceType is invalid</exception>
    private Task<string> AddCrossReferenceAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var referenceType = ArgumentHelper.GetString(arguments, "referenceType");
            var referenceText = ArgumentHelper.GetStringNullable(arguments, "referenceText");
            var targetName = ArgumentHelper.GetString(arguments, "targetName");
            var insertAsHyperlink = ArgumentHelper.GetBool(arguments, "insertAsHyperlink", true);
            var includeAboveBelow = ArgumentHelper.GetBool(arguments, "includeAboveBelow", false);

            var validTypes = new[] { "Heading", "Bookmark", "Figure", "Table", "Equation" };
            if (!validTypes.Contains(referenceType, StringComparer.OrdinalIgnoreCase))
                throw new ArgumentException(
                    $"Invalid referenceType: {referenceType}. Valid types are: {string.Join(", ", validTypes)}");

            var doc = new Document(path);
            var builder = new DocumentBuilder(doc);
            builder.MoveToDocumentEnd();

            if (!string.IsNullOrEmpty(referenceText))
                builder.Write(referenceText);

            var fieldCode = insertAsHyperlink ? $"REF {targetName} \\h" : $"REF {targetName}";

            builder.InsertField(fieldCode);
            if (includeAboveBelow)
                builder.Write(" (above)");

            doc.Save(outputPath);
            return $"Cross-reference added (Type: {referenceType}): {outputPath}";
        });
    }
}