using System.Text.Json.Nodes;
using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Word.Format;

/// <summary>
///     Handler for adding tab stops in Word documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class AddTabStopWordHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "add_tab";

    /// <summary>
    ///     Adds a tab stop to a paragraph.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: paragraphIndex, tabPosition
    ///     Optional: tabAlignment, tabLeader
    /// </param>
    /// <returns>Success message.</returns>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractAddTabStopParameters(parameters);

        var doc = context.Document;
        var para = ParagraphResolver.Resolve(doc, ParagraphAddress.From(parameters, p.ParagraphIndex)).Paragraph;

        // The same names, read the same way, as the array form. Spelling one wrong used to give
        // a left stop with no leader and a success message quoting the spelling back (R13-W01).
        var stop = WordTabStopHelper.Resolve(new JsonArray(new JsonObject
        {
            ["position"] = p.TabPosition,
            ["alignment"] = p.TabAlignment,
            ["leader"] = p.TabLeader
        }))[0];

        para.ParagraphFormat.TabStops.Add(stop);

        MarkModified(context);
        return new SuccessResult { Message = $"Tab stop added at {p.TabPosition}pt ({p.TabAlignment}, {p.TabLeader})" };
    }

    private static AddTabStopParameters ExtractAddTabStopParameters(OperationParameters parameters)
    {
        return new AddTabStopParameters(
            parameters.GetOptional("paragraphIndex", 0),
            parameters.GetOptional("tabPosition", 0.0),
            parameters.GetOptional("tabAlignment", "left"),
            parameters.GetOptional("tabLeader", "none"));
    }

    private sealed record AddTabStopParameters(
        int ParagraphIndex,
        double TabPosition,
        string TabAlignment,
        string TabLeader);
}
