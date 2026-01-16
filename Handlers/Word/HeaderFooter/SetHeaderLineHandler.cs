using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using WordParagraph = Aspose.Words.Paragraph;
using Section = Aspose.Words.Section;

namespace AsposeMcpServer.Handlers.Word.HeaderFooter;

/// <summary>
///     Handler for setting header lines in Word documents.
/// </summary>
public class SetHeaderLineHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "set_header_line";

    /// <summary>
    ///     Sets a line separator in the document header.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: lineStyle, lineWidth, sectionIndex, headerFooterType
    /// </param>
    /// <returns>Success message.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractSetHeaderLineParameters(parameters);

        var doc = context.Document;
        var hfType = WordHeaderFooterHelper.GetHeaderFooterType(p.HeaderFooterType, true);
        var sections = p.SectionIndex == -1 ? doc.Sections.Cast<Section>() : [doc.Sections[p.SectionIndex]];

        foreach (var section in sections)
        {
            var header = WordHeaderFooterHelper.GetOrCreateHeaderFooter(section, doc, hfType);

            var para = new WordParagraph(doc);
            para.ParagraphFormat.Borders.Bottom.LineStyle = p.LineStyle.ToLower() switch
            {
                "double" => LineStyle.Double,
                "thick" => LineStyle.Thick,
                _ => LineStyle.Single
            };

            if (p.LineWidth.HasValue) para.ParagraphFormat.Borders.Bottom.LineWidth = p.LineWidth.Value;

            header.AppendChild(para);
        }

        MarkModified(context);

        return Success("Header line set");
    }

    /// <summary>
    ///     Extracts parameters for the set header line operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static SetHeaderLineParameters ExtractSetHeaderLineParameters(OperationParameters parameters)
    {
        return new SetHeaderLineParameters(
            parameters.GetOptional("lineStyle", "single"),
            parameters.GetOptional<double?>("lineWidth"),
            parameters.GetOptional("sectionIndex", 0),
            parameters.GetOptional("headerFooterType", "primary")
        );
    }

    /// <summary>
    ///     Parameters for the set header line operation.
    /// </summary>
    /// <param name="LineStyle">The line style.</param>
    /// <param name="LineWidth">The line width.</param>
    /// <param name="SectionIndex">The section index.</param>
    /// <param name="HeaderFooterType">The header/footer type.</param>
    private sealed record SetHeaderLineParameters(
        string LineStyle,
        double? LineWidth,
        int SectionIndex,
        string HeaderFooterType);
}
