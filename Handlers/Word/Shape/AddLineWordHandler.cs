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
        var lineParams = ExtractLineParameters(parameters);
        var doc = context.Document;
        var section = doc.FirstSection;
        var calculatedWidth = lineParams.Width ?? CalculateDefaultWidth(section);

        var (targetNode, locationDesc) = GetTargetNode(doc, section, lineParams.Location);
        if (targetNode == null)
            throw new InvalidOperationException($"Could not access {lineParams.Location}");

        var linePara = CreateLineParagraph(doc, lineParams, calculatedWidth);
        InsertParagraph(targetNode, linePara, lineParams.Position);

        MarkModified(context);
        return $"Successfully inserted line in {locationDesc} at {lineParams.Position} position.";
    }

    /// <summary>
    ///     Extracts line parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted line parameters.</returns>
    private static LineParameters ExtractLineParameters(OperationParameters parameters)
    {
        return new LineParameters(
            parameters.GetOptional("location", "body"),
            parameters.GetOptional("position", "end"),
            parameters.GetOptional("lineStyle", "shape"),
            parameters.GetOptional("lineWidth", 1.0),
            parameters.GetOptional("lineColor", "000000"),
            parameters.GetOptional<double?>("width")
        );
    }

    /// <summary>
    ///     Calculates the default line width based on page setup.
    /// </summary>
    /// <param name="section">The document section.</param>
    /// <returns>The calculated width.</returns>
    private static double CalculateDefaultWidth(Section section)
    {
        return section.PageSetup.PageWidth - section.PageSetup.LeftMargin - section.PageSetup.RightMargin;
    }

    /// <summary>
    ///     Gets the target node based on location.
    /// </summary>
    /// <param name="doc">The Word document.</param>
    /// <param name="section">The document section.</param>
    /// <param name="location">The location string.</param>
    /// <returns>A tuple containing the target node and location description.</returns>
    private static (Node? targetNode, string locationDesc) GetTargetNode(Document doc, Section section, string location)
    {
        return location.ToLower() switch
        {
            "header" => (GetOrCreateHeaderFooter(doc, section, HeaderFooterType.HeaderPrimary), "header"),
            "footer" => (GetOrCreateHeaderFooter(doc, section, HeaderFooterType.FooterPrimary), "footer"),
            _ => (section.Body, "document body")
        };
    }

    /// <summary>
    ///     Gets or creates a header/footer of the specified type.
    /// </summary>
    /// <param name="doc">The Word document.</param>
    /// <param name="section">The document section.</param>
    /// <param name="type">The header/footer type.</param>
    /// <returns>The header/footer.</returns>
    private static WordHeaderFooter GetOrCreateHeaderFooter(Document doc, Section section, HeaderFooterType type)
    {
        var headerFooter = section.HeadersFooters[type];
        if (headerFooter == null)
        {
            headerFooter = new WordHeaderFooter(doc, type);
            section.HeadersFooters.Add(headerFooter);
        }

        return headerFooter;
    }

    /// <summary>
    ///     Creates a paragraph containing the line.
    /// </summary>
    /// <param name="doc">The Word document.</param>
    /// <param name="p">The line parameters.</param>
    /// <param name="width">The line width.</param>
    /// <returns>The paragraph containing the line.</returns>
    private static WordParagraph CreateLineParagraph(Document doc, LineParameters p, double width)
    {
        return p.LineStyle == "shape"
            ? CreateShapeLineParagraph(doc, p, width)
            : CreateBorderLineParagraph(doc, p);
    }

    /// <summary>
    ///     Creates a paragraph with a shape line.
    /// </summary>
    /// <param name="doc">The Word document.</param>
    /// <param name="p">The line parameters.</param>
    /// <param name="width">The line width.</param>
    /// <returns>The paragraph containing the shape line.</returns>
    private static WordParagraph CreateShapeLineParagraph(Document doc, LineParameters p, double width)
    {
        var linePara = new WordParagraph(doc)
        {
            ParagraphFormat =
                { SpaceBefore = 0, SpaceAfter = 0, LineSpacing = 1, LineSpacingRule = LineSpacingRule.Exactly }
        };

        var shape = new Aspose.Words.Drawing.Shape(doc, ShapeType.Line)
        {
            Width = width, Height = 0, StrokeWeight = p.LineWidth,
            StrokeColor = ColorHelper.ParseColor(p.LineColor), WrapType = WrapType.Inline
        };

        linePara.AppendChild(shape);
        return linePara;
    }

    /// <summary>
    ///     Creates a paragraph with a border line.
    /// </summary>
    /// <param name="doc">The Word document.</param>
    /// <param name="p">The line parameters.</param>
    /// <returns>The paragraph containing the border line.</returns>
    private static WordParagraph CreateBorderLineParagraph(Document doc, LineParameters p)
    {
        var linePara = new WordParagraph(doc);
        linePara.ParagraphFormat.Borders.Bottom.LineStyle = LineStyle.Single;
        linePara.ParagraphFormat.Borders.Bottom.LineWidth = p.LineWidth;
        linePara.ParagraphFormat.Borders.Bottom.Color = ColorHelper.ParseColor(p.LineColor);
        return linePara;
    }

    /// <summary>
    ///     Inserts the line paragraph at the target position.
    /// </summary>
    /// <param name="targetNode">The target node.</param>
    /// <param name="linePara">The line paragraph.</param>
    /// <param name="position">The position (start or end).</param>
    private static void InsertParagraph(Node targetNode, WordParagraph linePara, string position)
    {
        if (position == "start")
        {
            if (targetNode is WordHeaderFooter hf) hf.PrependChild(linePara);
            else if (targetNode is Body body) body.PrependChild(linePara);
        }
        else
        {
            if (targetNode is WordHeaderFooter hf) hf.AppendChild(linePara);
            else if (targetNode is Body body) body.AppendChild(linePara);
        }
    }

    /// <summary>
    ///     Record to hold line creation parameters.
    /// </summary>
    private sealed record LineParameters(
        string Location,
        string Position,
        string LineStyle,
        double LineWidth,
        string LineColor,
        double? Width);
}
