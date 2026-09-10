using System.Text.Json.Nodes;
using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Word.Reference;

/// <summary>
///     Handler for adding index entries to Word documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class AddIndexWordHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "add_index";

    /// <summary>
    ///     Adds index entries and optionally an INDEX field to the document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: indexEntries (JSON array)
    ///     Optional: insertIndexAtEnd (default: true), headingStyle (default: Heading 1)
    /// </param>
    /// <returns>Success message naming the entries inserted and those skipped.</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when the JSON is not an array, or holds more entries than one call may insert.
    /// </exception>
    public override object
        Execute(OperationContext<Document> context,
            OperationParameters parameters)
    {
        var p = ExtractAddIndexParameters(parameters);

        var indexEntriesArray = JsonNode.Parse(p.IndexEntriesJson)?.AsArray()
                                ?? throw new ArgumentException("indexEntries must be a valid JSON array");

        // The array arrives as a JSON string, so the bound applied to array parameters at the tool
        // boundary never saw these elements: the count only exists after decoding. Checking it here
        // is the first point at which it can be checked, and it happens before the first field is
        // inserted rather than after some of them are (R3-R05).
        SecurityHelper.ValidateArraySize(indexEntriesArray, "indexEntries");

        var doc = context.Document;
        var builder = new DocumentBuilder(doc);
        var inserted = 0;
        var skipped = 0;

        foreach (var entryObj in indexEntriesArray)
        {
            // An element that is not an object, or carries no usable text, produces no field. The
            // result used to report the length of the array, so those entries were reported as
            // added while nothing was written for them (R3-C06).
            if (entryObj is not JsonObject entry)
            {
                skipped++;
                continue;
            }

            var text = ReadString(entry, "text");
            if (string.IsNullOrEmpty(text))
            {
                skipped++;
                continue;
            }

            var subEntry = ReadString(entry, "subEntry");
            var pageRangeBookmark = ReadString(entry, "pageRangeBookmark");

            builder.MoveToDocumentEnd();
            var xeField = $"XE \"{text}\"";
            if (!string.IsNullOrEmpty(subEntry))
                xeField += $" \\t \"{subEntry}\"";
            if (!string.IsNullOrEmpty(pageRangeBookmark))
                xeField += $" \\r \"{pageRangeBookmark}\"";
            builder.InsertField(xeField);
            inserted++;
        }

        if (p.InsertIndexAtEnd)
        {
            builder.MoveToDocumentEnd();
            builder.InsertBreak(BreakType.PageBreak);

            var style = doc.Styles[p.HeadingStyle];
            if (style == null)
                builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            else
                builder.ParagraphFormat.Style = style;

            builder.Writeln("Index");
            builder.ParagraphFormat.Style = doc.Styles["Normal"];
            builder.InsertField("INDEX \\e \" \" \\h \"A\"");
        }

        MarkModified(context);

        var message = $"Index entries added. Inserted: {inserted}";
        if (skipped > 0) message += $" (skipped {skipped} entries with no usable text)";

        return new SuccessResult { Message = message };
    }

    /// <summary>
    ///     Reads a property as text, treating a non-string value as absent.
    /// </summary>
    /// <param name="entry">The entry object.</param>
    /// <param name="name">Property to read.</param>
    /// <returns>The value, or <c>null</c> when the property is missing or not a string.</returns>
    private static string? ReadString(JsonObject entry, string name)
    {
        return entry[name] is JsonValue value && value.TryGetValue<string>(out var text) ? text : null;
    }

    private static AddIndexParameters ExtractAddIndexParameters(OperationParameters parameters)
    {
        var indexEntriesJson = parameters.GetOptional<string?>("indexEntries");

        if (string.IsNullOrEmpty(indexEntriesJson))
            throw new ArgumentException("indexEntries is required for add_index operation");

        return new AddIndexParameters(
            indexEntriesJson,
            parameters.GetOptional("insertIndexAtEnd", true),
            parameters.GetOptional("headingStyle", "Heading 1"));
    }

    private sealed record AddIndexParameters(
        string IndexEntriesJson,
        bool InsertIndexAtEnd,
        string HeadingStyle);
}
