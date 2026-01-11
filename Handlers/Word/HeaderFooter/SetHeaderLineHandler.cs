using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using WordParagraph = Aspose.Words.Paragraph;
using Section = Aspose.Words.Section;

namespace AsposeMcpServer.Handlers.Word.HeaderFooter;

public class SetHeaderLineHandler : OperationHandlerBase<Document>
{
    public override string Operation => "set_header_line";

    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var lineStyle = parameters.GetOptional("lineStyle", "single");
        var lineWidth = parameters.GetOptional<double?>("lineWidth");
        var sectionIndex = parameters.GetOptional("sectionIndex", 0);
        var headerFooterType = parameters.GetOptional("headerFooterType", "primary");

        var doc = context.Document;
        var hfType = WordHeaderFooterHelper.GetHeaderFooterType(headerFooterType, true);
        var sections = sectionIndex == -1 ? doc.Sections.Cast<Section>() : [doc.Sections[sectionIndex]];

        foreach (var section in sections)
        {
            var header = WordHeaderFooterHelper.GetOrCreateHeaderFooter(section, doc, hfType);

            var para = new WordParagraph(doc);
            para.ParagraphFormat.Borders.Bottom.LineStyle = lineStyle.ToLower() switch
            {
                "double" => LineStyle.Double,
                "thick" => LineStyle.Thick,
                _ => LineStyle.Single
            };

            if (lineWidth.HasValue) para.ParagraphFormat.Borders.Bottom.LineWidth = lineWidth.Value;

            header.AppendChild(para);
        }

        MarkModified(context);

        return Success("Header line set");
    }
}
