using Aspose.Words;
using Aspose.Words.Drawing;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;
using WordParagraph = Aspose.Words.Paragraph;
using WordHeaderFooter = Aspose.Words.HeaderFooter;

namespace AsposeMcpServer.Handlers.Word.Shape;

/// <summary>
///     Handler for adding line shapes to Word documents.
/// </summary>
public class AddLineWordHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "add_line";

    /// <summary>
    ///     Adds a line shape to the document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: location, position, lineStyle, lineWidth, lineColor, width
    /// </param>
    /// <returns>Success message with line details.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var location = parameters.GetOptional("location", "body");
        var position = parameters.GetOptional("position", "end");
        var lineStyle = parameters.GetOptional("lineStyle", "shape");
        var lineWidth = parameters.GetOptional("lineWidth", 1.0);
        var lineColor = parameters.GetOptional("lineColor", "000000");
        var width = parameters.GetOptional<double?>("width");

        var doc = context.Document;
        var section = doc.FirstSection;
        var calculatedWidth = width ?? section.PageSetup.PageWidth - section.PageSetup.LeftMargin -
            section.PageSetup.RightMargin;

        Node? targetNode;
        string locationDesc;

        switch (location.ToLower())
        {
            case "header":
                var header = section.HeadersFooters[HeaderFooterType.HeaderPrimary];
                if (header == null)
                {
                    header = new WordHeaderFooter(doc, HeaderFooterType.HeaderPrimary);
                    section.HeadersFooters.Add(header);
                }

                targetNode = header;
                locationDesc = "header";
                break;

            case "footer":
                var footer = section.HeadersFooters[HeaderFooterType.FooterPrimary];
                if (footer == null)
                {
                    footer = new WordHeaderFooter(doc, HeaderFooterType.FooterPrimary);
                    section.HeadersFooters.Add(footer);
                }

                targetNode = footer;
                locationDesc = "footer";
                break;

            default:
                targetNode = section.Body;
                locationDesc = "document body";
                break;
        }

        if (targetNode == null)
            throw new InvalidOperationException($"Could not access {location}");

        if (lineStyle == "shape")
        {
            var linePara = new WordParagraph(doc)
            {
                ParagraphFormat =
                {
                    SpaceBefore = 0,
                    SpaceAfter = 0,
                    LineSpacing = 1,
                    LineSpacingRule = LineSpacingRule.Exactly
                }
            };

            var shape = new Aspose.Words.Drawing.Shape(doc, ShapeType.Line)
            {
                Width = calculatedWidth,
                Height = 0,
                StrokeWeight = lineWidth,
                StrokeColor = ColorHelper.ParseColor(lineColor),
                WrapType = WrapType.Inline
            };

            linePara.AppendChild(shape);

            if (position == "start")
            {
                if (targetNode is WordHeaderFooter headerFooter)
                    headerFooter.PrependChild(linePara);
                else if (targetNode is Body body)
                    body.PrependChild(linePara);
            }
            else
            {
                if (targetNode is WordHeaderFooter headerFooter)
                    headerFooter.AppendChild(linePara);
                else if (targetNode is Body body)
                    body.AppendChild(linePara);
            }
        }
        else
        {
            var linePara = new WordParagraph(doc);
            linePara.ParagraphFormat.Borders.Bottom.LineStyle = LineStyle.Single;
            linePara.ParagraphFormat.Borders.Bottom.LineWidth = lineWidth;
            linePara.ParagraphFormat.Borders.Bottom.Color = ColorHelper.ParseColor(lineColor);

            if (position == "start")
            {
                if (targetNode is WordHeaderFooter headerFooter)
                    headerFooter.PrependChild(linePara);
                else if (targetNode is Body body)
                    body.PrependChild(linePara);
            }
            else
            {
                if (targetNode is WordHeaderFooter headerFooter)
                    headerFooter.AppendChild(linePara);
                else if (targetNode is Body body)
                    body.AppendChild(linePara);
            }
        }

        MarkModified(context);

        return $"Successfully inserted line in {locationDesc} at {position} position.";
    }
}
