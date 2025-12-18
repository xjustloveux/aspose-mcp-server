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
- Add index: word_reference(operation='add_index', path='doc.docx', entries=[{'text':'Index term','page':1}])
- Add cross-reference: word_reference(operation='add_cross_reference', path='doc.docx', referenceType='Heading', targetText='Chapter 1', displayText='See Chapter 1')";

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

        return operation.ToLower() switch
        {
            "add_table_of_contents" => await AddTableOfContents(arguments),
            "update_table_of_contents" => await UpdateTableOfContents(arguments),
            "add_index" => await AddIndex(arguments),
            "add_cross_reference" => await AddCrossReference(arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds a table of contents to the document
    /// </summary>
    /// <param name="arguments">JSON arguments containing path, optional outputPath, headingLevels</param>
    /// <returns>Success message</returns>
    private async Task<string> AddTableOfContents(JsonObject? arguments)
    {
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");
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

        var switches = new List<string>
        {
            $"\\o \"1-{maxLevel}\""
        };

        if (!hyperlinks)
            switches.Add("\\n");

        if (!pageNumbers)
            switches.Add("\\n");

        if (!rightAlignPageNumbers)
            switches.Add("\\l");

        var fieldCode = $"TOC {string.Join(" ", switches)}";
        builder.InsertField(fieldCode);

        doc.Save(outputPath);
        return await Task.FromResult($"Table of contents added: {outputPath}");
    }

    /// <summary>
    ///     Updates the table of contents
    /// </summary>
    /// <param name="arguments">JSON arguments containing path, optional outputPath</param>
    /// <returns>Success message</returns>
    private async Task<string> UpdateTableOfContents(JsonObject? arguments)
    {
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");
        var tocIndex = ArgumentHelper.GetIntNullable(arguments, "tocIndex");

        var doc = new Document(path);
        // Search for TOC fields in the entire document (including headers/footers)
        var tocFields = doc.Range.Fields
            .Where(f => f.Type == FieldType.FieldTOC)
            .ToList();

        if (tocFields.Count == 0)
        {
            // Provide more helpful error message
            var allFields = doc.Range.Fields.ToList();
            var fieldTypes = allFields.Select(f => f.Type.ToString()).Distinct().ToList();
            var message = "No table of contents fields found in document.";
            if (allFields.Count > 0)
                message += $" Found {allFields.Count} field(s) of other types: {string.Join(", ", fieldTypes)}.";
            message += " Use 'add_table_of_contents' operation to add a table of contents first.";
            return await Task.FromResult(message);
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
        return await Task.FromResult($"Updated {updatedCount} table of contents field(s): {outputPath}");
    }

    /// <summary>
    ///     Adds an index to the document
    /// </summary>
    /// <param name="arguments">JSON arguments containing path, optional outputPath, entries</param>
    /// <returns>Success message</returns>
    private async Task<string> AddIndex(JsonObject? arguments)
    {
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");
        var indexEntriesArray = ArgumentHelper.GetArray(arguments, "indexEntries");
        var insertIndexAtEnd = ArgumentHelper.GetBool(arguments, "insertIndexAtEnd");
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
            builder.ParagraphFormat.Style = doc.Styles[headingStyle];
            builder.Writeln("Index");
            builder.ParagraphFormat.Style = doc.Styles["Normal"];
            builder.InsertField("INDEX \\e \" \" \\h \"A\"");
        }

        doc.Save(outputPath);
        return await Task.FromResult(
            $"Index entries added. Total entries: {indexEntriesArray.Count}. Output: {outputPath}");
    }

    /// <summary>
    ///     Adds a cross-reference to the document
    /// </summary>
    /// <param name="arguments">JSON arguments containing path, referenceType, target, optional outputPath</param>
    /// <returns>Success message</returns>
    private async Task<string> AddCrossReference(JsonObject? arguments)
    {
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        _ = ArgumentHelper.GetString(arguments, "referenceType");
        var referenceText = ArgumentHelper.GetStringNullable(arguments, "referenceText");
        var targetName = ArgumentHelper.GetString(arguments, "targetName");
        _ = ArgumentHelper.GetBool(arguments, "insertAsHyperlink");
        var includeAboveBelow = ArgumentHelper.GetBool(arguments, "includeAboveBelow", false);

        var doc = new Document(path);
        var builder = new DocumentBuilder(doc);
        builder.MoveToDocumentEnd();

        if (!string.IsNullOrEmpty(referenceText))
            builder.Write(referenceText);

        builder.InsertField($"REF {targetName} \\h");
        if (includeAboveBelow)
            builder.Write(" (above)");

        doc.Save(outputPath);
        return await Task.FromResult($"Cross-reference added: {outputPath}");
    }
}