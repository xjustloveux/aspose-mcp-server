using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Word.Reference;

/// <summary>
///     Handler for adding table of contents to Word documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class AddTableOfContentsWordHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "add_toc";

    /// <summary>
    ///     Adds a table of contents to the document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: position (default: start), title, maxLevel (default: 3), hyperlinks (default: true),
    ///     pageNumbers (default: true), rightAlignPageNumbers (default: true)
    /// </param>
    /// <returns>Success message indicating TOC was added.</returns>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractAddTableOfContentsParameters(parameters);

        var doc = context.Document;

        // The refusal comes first, while the document is still untouched. The title and the TOC
        // field were written before UpdateAllowedFields ran, so a document already holding a
        // nested disallowed field was refused with a heading and an empty table of contents
        // already added to it — and a session keeps that document (R8-W02).
        WordFieldPolicy.RefuseNestedDisallowedFields(doc, null);

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
        WordFieldPolicy.UpdateAllowedFields(doc);

        MarkModified(context);

        return new SuccessResult { Message = "Table of contents added" };
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

    private sealed record AddTableOfContentsParameters(
        string Position,
        string Title,
        int MaxLevel,
        bool Hyperlinks,
        bool PageNumbers,
        bool RightAlignPageNumbers);
}
