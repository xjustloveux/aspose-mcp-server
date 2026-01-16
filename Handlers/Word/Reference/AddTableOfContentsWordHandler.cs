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
        var p = ExtractAddTableOfContentsParameters(parameters);

        var doc = context.Document;
        var builder = new DocumentBuilder(doc);

        if (p.Position == "end")
            builder.MoveToDocumentEnd();
        else
            builder.MoveToDocumentStart();

        if (!string.IsNullOrEmpty(p.Title))
        {
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln(p.Title);
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        }

        var switches = $"\\o \"1-{p.MaxLevel}\"";

        if (!p.Hyperlinks)
            switches += " \\n";

        if (!p.PageNumbers)
            switches += " \\p \"\"";

        if (!p.RightAlignPageNumbers)
            switches += " \\l";

        builder.InsertTableOfContents(switches);
        doc.UpdateFields();

        MarkModified(context);

        return Success("Table of contents added");
    }

    private static AddTableOfContentsParameters ExtractAddTableOfContentsParameters(OperationParameters parameters)
    {
        return new AddTableOfContentsParameters(
            parameters.GetOptional("position", "start"),
            parameters.GetOptional("title", "Table of Contents"),
            parameters.GetOptional("maxLevel", 3),
            parameters.GetOptional("hyperlinks", true),
            parameters.GetOptional("pageNumbers", true),
            parameters.GetOptional("rightAlignPageNumbers", true));
    }

    private record AddTableOfContentsParameters(
        string Position,
        string Title,
        int MaxLevel,
        bool Hyperlinks,
        bool PageNumbers,
        bool RightAlignPageNumbers);
}
