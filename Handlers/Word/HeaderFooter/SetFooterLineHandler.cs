using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Results.Common;
using WordParagraph = Aspose.Words.Paragraph;
using Section = Aspose.Words.Section;

namespace AsposeMcpServer.Handlers.Word.HeaderFooter;

/// <summary>
///     Handler for setting footer lines in Word documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class SetFooterLineHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "set_footer_line";

    /// <summary>
    ///     Sets a line separator in the document footer.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: lineStyle, lineWidth, sectionIndex, headerFooterType
    /// </param>
    /// <returns>Success message.</returns>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractSetFooterLineParameters(parameters);

        var doc = context.Document;
        var hfType = WordHeaderFooterHelper.GetHeaderFooterType(p.HeaderFooterType, false);
        var sections = p.SectionIndex == -1 ? doc.Sections.Cast<Section>() : [doc.Sections[p.SectionIndex]];

        foreach (var section in sections)
        {
            var footer = WordHeaderFooterHelper.GetOrCreateHeaderFooter(section, doc, hfType);

            var para = new WordParagraph(doc);
            para.ParagraphFormat.Borders.Top.LineStyle = p.LineStyle.ToLower() switch
            {
                "double" => LineStyle.Double,
                "thick" => LineStyle.Thick,
                _ => LineStyle.Single
            };

            if (p.LineWidth.HasValue) para.ParagraphFormat.Borders.Top.LineWidth = p.LineWidth.Value;

            footer.AppendChild(para);
        }

        MarkModified(context);

        return new SuccessResult { Message = "Footer line set" };
    }

    /// <summary>
    ///     Extracts parameters for the set footer line operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static SetFooterLineParameters ExtractSetFooterLineParameters(OperationParameters parameters)
    {
        return new SetFooterLineParameters(
            parameters.GetOptional("lineStyle", "single"),
            parameters.GetOptional<double?>("lineWidth"),
            parameters.GetOptional("sectionIndex", 0),
            parameters.GetOptional("headerFooterType", "primary")
        );
    }

    /// <summary>
    ///     Parameters for the set footer line operation.
    /// </summary>
    /// <param name="LineStyle">The line style.</param>
    /// <param name="LineWidth">The line width.</param>
    /// <param name="SectionIndex">The section index.</param>
    /// <param name="HeaderFooterType">The header/footer type.</param>
    private sealed record SetFooterLineParameters(
        string LineStyle,
        double? LineWidth,
        int SectionIndex,
        string HeaderFooterType);
}
