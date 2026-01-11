using Aspose.Words;
using Aspose.Words.Drawing;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;
using WordParagraph = Aspose.Words.Paragraph;

namespace AsposeMcpServer.Handlers.Word.Shape;

/// <summary>
///     Handler for adding textboxes to Word documents.
/// </summary>
public class AddTextBoxWordHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "add_textbox";

    /// <summary>
    ///     Adds a textbox to the document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: text
    ///     Optional: textboxWidth, textboxHeight, positionX, positionY, backgroundColor, borderColor,
    ///     borderWidth, fontName, fontNameAscii, fontNameFarEast, fontSize, bold, textAlignment
    /// </param>
    /// <returns>Success message.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var text = parameters.GetOptional<string?>("text");
        var textboxWidth = parameters.GetOptional("textboxWidth", 200.0);
        var textboxHeight = parameters.GetOptional("textboxHeight", 100.0);
        var positionX = parameters.GetOptional("positionX", 100.0);
        var positionY = parameters.GetOptional("positionY", 100.0);
        var backgroundColor = parameters.GetOptional<string?>("backgroundColor");
        var borderColor = parameters.GetOptional<string?>("borderColor");
        var borderWidth = parameters.GetOptional("borderWidth", 1.0);
        var fontName = parameters.GetOptional<string?>("fontName");
        var fontNameAscii = parameters.GetOptional<string?>("fontNameAscii");
        var fontNameFarEast = parameters.GetOptional<string?>("fontNameFarEast");
        var fontSize = parameters.GetOptional<double?>("fontSize");
        var bold = parameters.GetOptional<bool?>("bold");
        var textAlignment = parameters.GetOptional("textAlignment", "left");

        if (string.IsNullOrEmpty(text))
            throw new ArgumentException("text is required for add_textbox operation");

        var doc = context.Document;
        var builder = new DocumentBuilder(doc);
        builder.MoveToDocumentEnd();

        var textBox = new Aspose.Words.Drawing.Shape(doc, ShapeType.TextBox)
        {
            Width = textboxWidth,
            Height = textboxHeight,
            Left = positionX,
            Top = positionY,
            WrapType = WrapType.None,
            RelativeHorizontalPosition = RelativeHorizontalPosition.Page,
            RelativeVerticalPosition = RelativeVerticalPosition.Page
        };

        if (!string.IsNullOrEmpty(backgroundColor))
        {
            textBox.Fill.Color = ColorHelper.ParseColor(backgroundColor);
            textBox.Fill.Visible = true;
        }

        if (!string.IsNullOrEmpty(borderColor))
        {
            textBox.Stroke.Color = ColorHelper.ParseColor(borderColor);
            textBox.Stroke.Weight = borderWidth;
            textBox.Stroke.Visible = true;
        }

        var para = new WordParagraph(doc);
        var run = new Run(doc, text);

        if (!string.IsNullOrEmpty(fontNameAscii))
            run.Font.NameAscii = fontNameAscii;

        if (!string.IsNullOrEmpty(fontNameFarEast))
            run.Font.NameFarEast = fontNameFarEast;

        if (!string.IsNullOrEmpty(fontName))
        {
            if (string.IsNullOrEmpty(fontNameAscii) && string.IsNullOrEmpty(fontNameFarEast))
            {
                run.Font.Name = fontName;
            }
            else
            {
                if (string.IsNullOrEmpty(fontNameAscii))
                    run.Font.NameAscii = fontName;
                if (string.IsNullOrEmpty(fontNameFarEast))
                    run.Font.NameFarEast = fontName;
            }
        }

        if (fontSize.HasValue)
            run.Font.Size = fontSize.Value;

        if (bold.HasValue)
            run.Font.Bold = bold.Value;

        para.ParagraphFormat.Alignment = WordShapeHelper.ParseAlignment(textAlignment);

        para.AppendChild(run);
        textBox.AppendChild(para);
        builder.InsertNode(textBox);

        MarkModified(context);

        return "Successfully added textbox.";
    }
}
