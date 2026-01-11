using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Reference;

/// <summary>
///     Handler for adding table of contents to Word documents.
/// </summary>
public class AddTableOfContentsWordHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "add_table_of_contents";

    /// <summary>
    ///     Adds a table of contents to the document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: position (default: start), title, maxLevel (default: 3), hyperlinks (default: true),
    ///     pageNumbers (default: true), rightAlignPageNumbers (default: true)
    /// </param>
    /// <returns>Success message indicating TOC was added.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var position = parameters.GetOptional("position", "start");
        var title = parameters.GetOptional("title", "Table of Contents");
        var maxLevel = parameters.GetOptional("maxLevel", 3);
        var hyperlinks = parameters.GetOptional("hyperlinks", true);
        var pageNumbers = parameters.GetOptional("pageNumbers", true);
        var rightAlignPageNumbers = parameters.GetOptional("rightAlignPageNumbers", true);

        var doc = context.Document;
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

        var switches = $"\\o \"1-{maxLevel}\"";

        if (!hyperlinks)
            switches += " \\n";

        if (!pageNumbers)
            switches += " \\p \"\"";

        if (!rightAlignPageNumbers)
            switches += " \\l";

        builder.InsertTableOfContents(switches);
        doc.UpdateFields();

        MarkModified(context);

        return Success("Table of contents added");
    }
}
