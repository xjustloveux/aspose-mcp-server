using System.Text.Json.Nodes;
using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Reference;

/// <summary>
///     Handler for adding index entries to Word documents.
/// </summary>
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
    /// <returns>Success message with count of entries added.</returns>
    public override string
        Execute(OperationContext<Document> context,
            OperationParameters parameters) // NOSONAR S3776 - Linear field code building
    {
        var p = ExtractAddIndexParameters(parameters);

        var indexEntriesArray = JsonNode.Parse(p.IndexEntriesJson)?.AsArray()
                                ?? throw new ArgumentException("indexEntries must be a valid JSON array");

        var doc = context.Document;
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

        return Success($"Index entries added. Total entries: {indexEntriesArray.Count}");
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
