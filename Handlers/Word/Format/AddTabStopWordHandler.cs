using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Format;

/// <summary>
///     Handler for adding tab stops in Word documents.
/// </summary>
public class AddTabStopWordHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "add_tab_stop";

    /// <summary>
    ///     Adds a tab stop to a paragraph.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: paragraphIndex, tabPosition
    ///     Optional: tabAlignment, tabLeader
    /// </param>
    /// <returns>Success message.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractAddTabStopParameters(parameters);

        var doc = context.Document;
        var para = WordFormatHelper.GetTargetParagraph(doc, p.ParagraphIndex);

        var tabAlignment = p.TabAlignment.ToLower() switch
        {
            "center" => TabAlignment.Center,
            "right" => TabAlignment.Right,
            "decimal" => TabAlignment.Decimal,
            "bar" => TabAlignment.Bar,
            _ => TabAlignment.Left
        };

        var tabLeader = p.TabLeader.ToLower() switch
        {
            "dots" => TabLeader.Dots,
            "dashes" => TabLeader.Dashes,
            "line" => TabLeader.Line,
            "heavy" => TabLeader.Heavy,
            "middledot" => TabLeader.MiddleDot,
            _ => TabLeader.None
        };

        para.ParagraphFormat.TabStops.Add(new TabStop(p.TabPosition, tabAlignment, tabLeader));

        MarkModified(context);
        return Success($"Tab stop added at {p.TabPosition}pt ({p.TabAlignment}, {p.TabLeader})");
    }

    private static AddTabStopParameters ExtractAddTabStopParameters(OperationParameters parameters)
    {
        return new AddTabStopParameters(
            parameters.GetOptional("paragraphIndex", 0),
            parameters.GetOptional("tabPosition", 0.0),
            parameters.GetOptional("tabAlignment", "left"),
            parameters.GetOptional("tabLeader", "none"));
    }

    private record AddTabStopParameters(
        int ParagraphIndex,
        double TabPosition,
        string TabAlignment,
        string TabLeader);
}
